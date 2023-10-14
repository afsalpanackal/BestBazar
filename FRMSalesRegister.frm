VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMSalesReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES & PURCHASE REPORT FOR E-FILING"
   ClientHeight    =   9480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   19305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMSalesRegister.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   19305
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   9765
      Left            =   -120
      TabIndex        =   0
      Top             =   -270
      Width           =   19365
      Begin VB.CommandButton CmdCrDr 
         Caption         =   "Credit / Debit Note GST Report"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   12660
         TabIndex        =   5
         Top             =   7665
         Width           =   1560
      End
      Begin VB.Frame Frame1 
         Caption         =   "3rd Rate Reports"
         Height          =   1770
         Left            =   16500
         TabIndex        =   48
         Top             =   7455
         Width           =   2820
         Begin VB.OptionButton OptComb1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "B2B + B2C (Combined)"
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
            Height          =   435
            Left            =   60
            TabIndex        =   63
            Top             =   1305
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton OptSeperate1 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Separate"
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
            Height          =   435
            Left            =   1440
            TabIndex        =   62
            Top             =   1305
            Width           =   1320
         End
         Begin VB.CommandButton CmdReset 
            Caption         =   "Reset ST Rate"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   1425
            TabIndex        =   52
            Top             =   765
            Width           =   1335
         End
         Begin VB.CommandButton Command8 
            Caption         =   "HSN Code Wise Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   510
            Left            =   45
            TabIndex        =   51
            Top             =   765
            Width           =   1335
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Sales-GSTR1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   1425
            TabIndex        =   50
            Top             =   255
            Width           =   1320
         End
         Begin VB.CommandButton CMD3RATE 
            Caption         =   "&Bill Wise Display With 3rd Rate"
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
            Left            =   45
            TabIndex        =   49
            Top             =   270
            Width           =   1335
         End
      End
      Begin VB.CommandButton CmdGST 
         Caption         =   "GST Reports"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   14250
         TabIndex        =   6
         Top             =   7665
         Width           =   1185
      End
      Begin VB.CommandButton CMDGSTR1 
         Caption         =   "Sales-  GSTR1"
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
         Left            =   11580
         TabIndex        =   4
         Top             =   7680
         Width           =   1035
      End
      Begin VB.Frame FramSALES 
         Caption         =   "Other Reports"
         Height          =   1425
         Left            =   165
         TabIndex        =   40
         Top             =   7455
         Width           =   8115
         Begin VB.CommandButton CmdJSON 
            Caption         =   "Generate JSON File"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   6645
            TabIndex        =   65
            Top             =   810
            Width           =   1410
         End
         Begin VB.CommandButton Command7 
            Caption         =   "HSN Code Wise Sales Summary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   2550
            TabIndex        =   58
            Top             =   225
            Width           =   1305
         End
         Begin VB.OptionButton optseperate 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Separate Report"
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   2265
            TabIndex        =   46
            Top             =   840
            Width           =   2025
         End
         Begin VB.OptionButton optcombine 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Combined Report"
            ForeColor       =   &H000000FF&
            Height          =   270
            Left            =   75
            TabIndex        =   45
            Top             =   840
            Value           =   -1  'True
            Width           =   2175
         End
         Begin VB.CommandButton Command5 
            Caption         =   "HSN Code Wise Purchase Report"
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
            Left            =   5235
            TabIndex        =   44
            Top             =   240
            Width           =   1380
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Print Item Wise Purchase Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   6645
            TabIndex        =   43
            Top             =   225
            Width           =   1410
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Print Item Wise Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   3915
            TabIndex        =   10
            Top             =   225
            Width           =   1245
         End
         Begin VB.CommandButton Command1 
            Caption         =   "HSN Code Wise Sales Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   1245
            TabIndex        =   9
            Top             =   225
            Width           =   1275
         End
         Begin VB.CommandButton Command4 
            Caption         =   "Print HSN (Bill Wise) Report"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   570
            Left            =   60
            TabIndex        =   8
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.CommandButton CmdSummary 
         Caption         =   "Day Wise Sales"
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
         Left            =   10500
         TabIndex        =   3
         Top             =   7680
         Width           =   1050
      End
      Begin VB.CommandButton CMDREGISTER 
         Caption         =   "EXPORT TO EXCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8295
         TabIndex        =   1
         Top             =   7695
         Width           =   1065
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   15450
         TabIndex        =   7
         Top             =   7680
         Width           =   1020
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&Bill Wise Display"
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
         Left            =   9375
         TabIndex        =   2
         Top             =   7680
         Width           =   1080
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   5325
         Left            =   150
         TabIndex        =   22
         Top             =   1320
         Width           =   19140
         _ExtentX        =   33761
         _ExtentY        =   9393
         _Version        =   393216
         Rows            =   1
         Cols            =   19
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
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
      Begin MSFlexGridLib.MSFlexGrid GrdTotal 
         Height          =   795
         Left            =   150
         TabIndex        =   29
         Top             =   6660
         Width           =   19140
         _ExtentX        =   33761
         _ExtentY        =   1402
         _Version        =   393216
         Rows            =   1
         Cols            =   18
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColor       =   16765606
         BackColorFixed  =   0
         ForeColorFixed  =   65535
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
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
         Height          =   1125
         Left            =   150
         TabIndex        =   11
         Top             =   210
         Width           =   19155
         Begin VB.Frame FrmPurchase 
            BackColor       =   &H00C0C0FF&
            Height          =   1110
            Left            =   9780
            TabIndex        =   34
            Top             =   15
            Width           =   3900
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0FF&
               Height          =   780
               Left            =   2205
               TabIndex        =   53
               Top             =   270
               Width           =   1650
               Begin VB.OptionButton OptRcvd 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Rcvd Date"
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
                  Height          =   270
                  Left            =   45
                  TabIndex        =   55
                  Top             =   450
                  Width           =   1320
               End
               Begin VB.OptionButton Optinvdate 
                  BackColor       =   &H00C0C0FF&
                  Caption         =   "Invoice Date"
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
                  Height          =   270
                  Left            =   30
                  TabIndex        =   54
                  Top             =   210
                  Value           =   -1  'True
                  Width           =   1560
               End
            End
            Begin VB.OptionButton OptComm 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Comm Purchase"
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
               Left            =   75
               TabIndex        =   37
               Top             =   735
               Visible         =   0   'False
               Width           =   2115
            End
            Begin VB.OptionButton OptLocal 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Local Purchase"
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
               Left            =   75
               TabIndex        =   36
               Top             =   420
               Width           =   2145
            End
            Begin VB.OptionButton OptNormal 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Purchase"
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
               Left            =   60
               TabIndex        =   35
               Top             =   135
               Value           =   -1  'True
               Width           =   1845
            End
         End
         Begin VB.Frame FrameSales 
            BackColor       =   &H00C0C0FF&
            Height          =   1110
            Left            =   13335
            TabIndex        =   19
            Top             =   15
            Visible         =   0   'False
            Width           =   5805
            Begin VB.OptionButton Optst 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Stock Transfer"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   2265
               TabIndex        =   61
               Top             =   390
               Width           =   1680
            End
            Begin VB.OptionButton Optbrsale 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Branch Sale"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   3945
               TabIndex        =   60
               Top             =   390
               Width           =   1650
            End
            Begin VB.Frame Frame4 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Branch wise "
               Height          =   555
               Left            =   2220
               TabIndex        =   59
               Top             =   165
               Width           =   3420
            End
            Begin VB.OptionButton OptComb 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Combined"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   3825
               TabIndex        =   47
               Top             =   870
               Width           =   1530
            End
            Begin VB.OptionButton OptService 
               BackColor       =   &H00C0C0FF&
               Caption         =   "GST - Service"
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
               Left            =   345
               TabIndex        =   41
               Top             =   795
               Width           =   1950
            End
            Begin VB.OptionButton OptGR 
               BackColor       =   &H00C0C0FF&
               Caption         =   "GST - B2C"
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
               Left            =   345
               TabIndex        =   39
               Top             =   480
               Value           =   -1  'True
               Width           =   1740
            End
            Begin VB.OptionButton OptGST 
               BackColor       =   &H00C0C0FF&
               Caption         =   "GST - B2B"
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
               Left            =   345
               TabIndex        =   38
               Top             =   165
               Width           =   1710
            End
            Begin VB.OptionButton Opt8V 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Other Sale"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   2300
               TabIndex        =   33
               Top             =   855
               Width           =   1530
            End
            Begin VB.OptionButton OptRT 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Old 8B Bills"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   2300
               TabIndex        =   21
               Top             =   1320
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.OptionButton OptWS 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Old 8 Bills"
               ForeColor       =   &H00FF0000&
               Height          =   210
               Left            =   2300
               TabIndex        =   20
               Top             =   1080
               Visible         =   0   'False
               Width           =   1800
            End
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   30
            TabIndex        =   12
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1590
            TabIndex        =   13
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   138149889
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   3420
            TabIndex        =   14
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   138149889
            CurrentDate     =   40498
         End
         Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
            Height          =   360
            Left            =   14850
            TabIndex        =   25
            Tag             =   "5"
            Top             =   480
            Width           =   4020
            _ExtentX        =   7091
            _ExtentY        =   635
            Picture         =   "FRMSalesRegister.frx":030A
            ForeColor       =   0
            BarPicture      =   "FRMSalesRegister.frx":0326
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
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0FF&
            Height          =   1110
            Left            =   4995
            TabIndex        =   26
            Top             =   15
            Width           =   4875
            Begin VB.OptionButton OptQtn 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Quotation Reg"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   2820
               TabIndex        =   66
               Top             =   840
               Width           =   1815
            End
            Begin VB.OptionButton OptDamage 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Damage from Customers"
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
               Height          =   270
               Left            =   50
               TabIndex        =   64
               Top             =   585
               Width           =   2490
            End
            Begin VB.OptionButton OptExpense 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Expense Bill"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   1380
               TabIndex        =   57
               Top             =   840
               Width           =   1815
            End
            Begin VB.OptionButton OptExReturn 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Exchange"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   45
               TabIndex        =   56
               Top             =   840
               Width           =   1815
            End
            Begin VB.OptionButton optAssets 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Assets Purchase"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   2535
               TabIndex        =   42
               Top             =   600
               Width           =   2085
            End
            Begin VB.OptionButton OptCST 
               BackColor       =   &H00C0C0FF&
               Caption         =   "CST Bill"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   2025
               TabIndex        =   32
               Top             =   1215
               Width           =   1260
            End
            Begin VB.OptionButton OptPurchret 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Purchase Return"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   2535
               TabIndex        =   31
               Top             =   360
               Width           =   2190
            End
            Begin VB.OptionButton OptSalesreturn 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Sales Return"
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   50
               TabIndex        =   30
               Top             =   375
               Width           =   1815
            End
            Begin VB.OptionButton optPurchase 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Purchase Bill"
               ForeColor       =   &H000000FF&
               Height          =   270
               Left            =   2535
               TabIndex        =   28
               Top             =   135
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton Optsales 
               BackColor       =   &H00C0C0FF&
               Caption         =   "Sales Bill"
               ForeColor       =   &H000000FF&
               Height          =   225
               Left            =   50
               TabIndex        =   27
               Top             =   135
               Width           =   1470
            End
         End
         Begin VB.Label LBLDEALER2 
            Height          =   315
            Left            =   0
            TabIndex        =   24
            Top             =   -345
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label FLAGCHANGE2 
            Height          =   315
            Left            =   0
            TabIndex        =   23
            Top             =   -360
            Width           =   495
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
            Left            =   1050
            TabIndex        =   18
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
            Left            =   3150
            TabIndex        =   17
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   16
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   15
            Top             =   1905
            Visible         =   0   'False
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "FRMSalesReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY_REC As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG, PHY_FLAG As Boolean
Dim Sum_flag As Boolean

Private Sub CMD3RATE_Click()
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim rstTRXMAST As ADODB.Recordset
    
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'db.Execute "UPDATE TRXFILE SET ST_RATE = 0 WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' "
    db.Execute "UPDATE ITEMMAST SET REMARKS = '' WHERE ISNULL(REMARKS) "
    db.Execute "UPDATE ITEMMAST SET P_VAN = 0 WHERE ISNULL(P_VAN) "
    db.Execute "UPDATE ITEMMAST SET P_RETAIL = 0 WHERE ISNULL(P_RETAIL) "
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(ST_RATE) OR ST_RATE = 0) ", db, adOpenStatic, adLockOptimistic, adCmdText
    rstTRANX.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rstTRANX.EOF
        Set rstTRXMAST = New ADODB.Recordset
        rstTRXMAST.Open "SELECT *  FROM  ITEMMAST WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        With rstTRXMAST
            If Not (.EOF And .BOF) Then
                If rstTRXMAST!P_VAN <> 0 And rstTRXMAST!P_VAN < rstTRANX!P_RETAIL Then
                    rstTRANX!ST_RATE = rstTRXMAST!P_VAN
                Else
                    rstTRANX!ST_RATE = rstTRANX!P_RETAIL
                End If
            End If
        End With
        rstTRXMAST.Close
        Set rstTRXMAST = Nothing
        
        rstTRANX.Update
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    db.Execute "Update TRXFILE SET ST_RATE = PTR WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(ST_RATE) OR ST_RATE = 0) "
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 9 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Sub
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            Dim KFC As Double
            TOTAL_AMT = 0
            KFC = 0
            Do Until n = GRDTranx.Cols - 3
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) - (((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 3)
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "KFC"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            
'            Set RSTtax = New ADODB.Recordset
'            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
'            Do Until RSTtax.EOF
'                GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL), 3)
'                RSTtax.MoveNext
'            Loop
'            RSTtax.Close
'            Set RSTtax = Nothing
            TOTAL_AMT = 0
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) - (((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100)
                End Select
                If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                    TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.Tag)
                    TOTAL_AMT = TOTAL_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                    CESSPER = CESSPER + ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                Else
                    TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                    TOTAL_AMT = TOTAL_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    CESSPER = CESSPER + ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) - ((RSTtax!ST_RATE * 100 / (RSTtax!SALES_TAX + 100)) * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                End If
                TOTAL_AMT = TOTAL_AMT + CESSPER + CESSAMT
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            TOTAL_AMT = Round(TOTAL_AMT, 2)
            
            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
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

Private Sub CmdCrDr_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim TAX_PER As Single
    Dim RSTACTMAST As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    On Error GoTo ERRHAND
    
    If MDIMAIN.lblgst.Caption = "R" Then
        GRDTranx.rows = 1
        GRDTranx.Cols = 14
        
        GRDTranx.TextMatrix(0, 0) = "GSTin No"
        GRDTranx.TextMatrix(0, 1) = "Receiver Name"
        GRDTranx.TextMatrix(0, 2) = "Invoice No."
        GRDTranx.TextMatrix(0, 3) = "Invoice Date"
        GRDTranx.TextMatrix(0, 4) = "Note No"
        GRDTranx.TextMatrix(0, 5) = "Note Date"
        GRDTranx.TextMatrix(0, 6) = "Document Type"
        GRDTranx.TextMatrix(0, 7) = "Place of Supply"
        GRDTranx.TextMatrix(0, 8) = "Amount"
        GRDTranx.TextMatrix(0, 9) = "Rate of Tax"
        GRDTranx.TextMatrix(0, 10) = "Applicable %"
        GRDTranx.TextMatrix(0, 11) = "Taxable value"
        GRDTranx.TextMatrix(0, 12) = "Cess Amount"
        GRDTranx.TextMatrix(0, 13) = "Pre GST"
        
        GRDTranx.ColWidth(0) = 1800
        GRDTranx.ColWidth(1) = 3000
        GRDTranx.ColWidth(2) = 1300
        GRDTranx.ColWidth(3) = 1400
        GRDTranx.ColWidth(4) = 1200
        GRDTranx.ColWidth(5) = 1400
        GRDTranx.ColWidth(6) = 1400
        GRDTranx.ColWidth(7) = 1700
        GRDTranx.ColWidth(8) = 1400
        GRDTranx.ColWidth(9) = 1400
        GRDTranx.ColWidth(10) = 1400
        GRDTranx.ColWidth(11) = 1400
        GRDTranx.ColWidth(12) = 1300
        GRDTranx.ColWidth(13) = 1200
        
        
        GRDTranx.ColAlignment(0) = 1
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 4
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GRDTranx.ColAlignment(5) = 4
        GRDTranx.ColAlignment(6) = 4
        GRDTranx.ColAlignment(7) = 4
        GRDTranx.ColAlignment(8) = 4
        GRDTranx.ColAlignment(9) = 4
        GRDTranx.ColAlignment(10) = 4
        GRDTranx.ColAlignment(11) = 4
        GRDTranx.ColAlignment(12) = 4
        GRDTranx.ColAlignment(13) = 4
        
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR')  ORDER BY TRX_TYPE, VCH_NO, VCH_DATE", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            
            Set RSTACTMAST = New ADODB.Recordset
            RSTACTMAST.Open "SELECT * FROM RETURNMAST WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & RSTACTMAST!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    GRDTranx.TextMatrix(M, 0) = IIf(IsNull(RSTTRXFILE!KGST), "", RSTTRXFILE!KGST)
                    If RSTTRXFILE!CUST_IGST = "Y" Then
                        GRDTranx.TextMatrix(M, 7) = "INTER STATE"
                    Else
                        GRDTranx.TextMatrix(M, 7) = "32-KERALA"
                    End If
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(RSTACTMAST!ACT_NAME), "", RSTACTMAST!ACT_NAME)
            End If
            RSTACTMAST.Close
            Set RSTACTMAST = Nothing
            
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!INV_DETAILS), "", rstTRANX!INV_DETAILS)
            GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!INV_DATE), rstTRANX!INV_DATE, "")
            GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!VCH_NO), "", "SR-" & Format(rstTRANX!VCH_NO, bill_for))
            '
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            GRDTranx.TextMatrix(M, 6) = "C"
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!TRX_TOTAL), 0, rstTRANX!TRX_TOTAL)
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!SALES_TAX), 0, rstTRANX!SALES_TAX)
            GRDTranx.TextMatrix(M, 10) = ""
            GRDTranx.TextMatrix(M, 11) = Format(Round((Val(GRDTranx.TextMatrix(M, 8)) * 100) / (Val(GRDTranx.TextMatrix(M, 9)) + 100), 3), "0.000")
            GRDTranx.TextMatrix(M, 12) = ""
            GRDTranx.TextMatrix(M, 13) = ""
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        
        '''debit note
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PR')  ORDER BY TRX_TYPE, VCH_NO, VCH_DATE", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            
            Set RSTACTMAST = New ADODB.Recordset
            RSTACTMAST.Open "SELECT * FROM PURCAHSERETURN WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & RSTACTMAST!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    GRDTranx.TextMatrix(M, 0) = IIf(IsNull(RSTTRXFILE!KGST), "", RSTTRXFILE!KGST)
                    If RSTTRXFILE!CUST_IGST = "Y" Then
                        GRDTranx.TextMatrix(M, 7) = "INTER STATE"
                    Else
                        GRDTranx.TextMatrix(M, 7) = "32-KERALA"
                    End If
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(RSTACTMAST!ACT_NAME), "", RSTACTMAST!ACT_NAME)
            End If
            RSTACTMAST.Close
            Set RSTACTMAST = Nothing
            
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!INV_DETAILS), "", rstTRANX!INV_DETAILS)
            GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!INV_DATE), rstTRANX!INV_DATE, "")
            GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!VCH_NO), "", Format(rstTRANX!VCH_NO, bill_for))
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            GRDTranx.TextMatrix(M, 6) = "D"
            GRDTranx.TextMatrix(M, 8) = IIf(IsNull(rstTRANX!TRX_TOTAL), 0, rstTRANX!TRX_TOTAL)
            GRDTranx.TextMatrix(M, 9) = IIf(IsNull(rstTRANX!SALES_TAX), 0, rstTRANX!SALES_TAX)
            GRDTranx.TextMatrix(M, 10) = ""
            GRDTranx.TextMatrix(M, 11) = Format(Round((Val(GRDTranx.TextMatrix(M, 8)) * 100) / (Val(GRDTranx.TextMatrix(M, 9)) + 100), 3), "0.000")
            GRDTranx.TextMatrix(M, 12) = ""
            GRDTranx.TextMatrix(M, 13) = ""
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmDDisplay_Click()
    Sum_flag = False
    Dim i, n As Long
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    If Optsales.Value = True Then
        'Call
        If Optbrsale.Value = True Then
            Call Sales_RegisterBR
        Else
            Call Sales_Register
        End If
    ElseIf OptPurchase.Value = True Then
        Call Purchase_Register
    ElseIf OptSalesreturn.Value = True Then
        Call SALES_RET_REGISTER
    ElseIf OptPurchret.Value = True Then
        Call PURCH_RET_Register
    ElseIf optAssets.Value = True Then
        Call Assets_Register(1)
    ElseIf OptExpense.Value = True Then
        Call Assets_Register(0)
    ElseIf OptExReturn.Value = True Then
        Call EX_RET_REGISTER
    ElseIf OptQtn.Value = True Then
        Call QTN_Register
    ElseIf OptDamage.Value = True Then
        Call Damage_Register
    Else
        Call Sales_Register_CST
    End If
    If Optsales.Value = True Or OptSalesreturn.Value = True Then
        
    Else
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        GrdTotal.TextMatrix(0, 3) = "TOTAL"
        For n = 4 To GRDTranx.Cols - 2
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTo.SetFocus
    End Select
End Sub

Private Sub CmdGST_Click()
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
    On Error GoTo ERRHAND
    'Create an Excel instalce.
    Set oApp = CreateObject("Excel.Application")
    
    Set oWB = oApp.Workbooks.Add
    Set oWS = oWB.Worksheets(1)
    
    oWS.Range("A" & 1).Value = "Invoice No."
    oWS.Range("B" & 1).Value = "Customer Name"
    oWS.Range("C" & 1).Value = "GSTIN"
    oWS.Range("D" & 1).Value = "Invoice Date"
    'oApp.Columns("D:D").NumberFormat = "dd-mmm-yy"
    oWS.Range("D:D").NumberFormat = "dd-mmm-yy"
    oWS.Range("E" & 1).Value = "Invoice Value"
    oWS.Range("F" & 1).Value = "Tax Rate(%)"
    oWS.Range("G" & 1).Value = "Taxable value"
    oWS.Range("H" & 1).Value = "IGST"
    oWS.Range("I" & 1).Value = "Central Tax"
    oWS.Range("J" & 1).Value = "State Tax"
    oWS.Range("K" & 1).Value = "Cess"
    oWS.Range("L" & 1).Value = "State of supply"
    oWS.Range("M" & 1).Value = "Reverse Charge"
    oWS.Range("N" & 1).Value = "E-com GSTIN"
    
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim M As Long
   
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
    Dim CESSPER As Double
    Dim CESSAMT As Double
    Dim TOTAL_AMT As Double
    
    On Error GoTo ERRHAND
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim rststock As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    M = 1
    Do Until rstTRANX.EOF
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            M = M + 1
            oWS.Range("A" & M).NumberFormat = "@"
            oWS.Range("A" & M).Value = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                oWS.Range("B" & M).Value = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            Else
                oWS.Range("B" & M).Value = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            End If
            oWS.Range("C" & M).Value = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            oWS.Range("D" & M).Value = IIf(IsDate(rstTRANX!VCH_DATE), Format(rstTRANX!VCH_DATE, "MM/DD/YYYY"), "")
            CESSPER = 0
            CESSAMT = 0
            TaxAmt = 0
            TOTAL_AMT = 0
            TAXSALEAMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTtax!SALES_TAX & "  AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) - ((RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100)
                End Select
                If IsNull(RSTTRXFILE!QTY) Or RSTTRXFILE!QTY = 0 Then
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100)
                    
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt)
                Else
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag) * Val(RSTTRXFILE!QTY)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100) * RSTTRXFILE!QTY
                    
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * RSTTRXFILE!QTY * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt) * RSTTRXFILE!QTY
                End If
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            oWS.Range("F" & M).Value = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
            oWS.Range("G" & M).Value = Round(TAXSALEAMT, 2)
            If rstTRANX!CUST_IGST = "Y" Or (Len(Trim(oWS.Range("C" & M).Value)) = 15 And Left(Trim(oWS.Range("C" & M).Value), 2) <> Trim(MDIMAIN.LBLSTATE.Caption)) Then
                oWS.Range("H" & M).Value = Round(TaxAmt, 2)
            Else
                oWS.Range("I" & M).Value = Round(TaxAmt / 2, 2)
                oWS.Range("J" & M).Value = Round(TaxAmt / 2, 2)
            End If
            oWS.Range("K" & M).Value = Round(CESSPER + CESSAMT, 2)
            DISC_AMT = 0
            'If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            'End If
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                TOTAL_AMT = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
            oWS.Range("E" & M).Value = Round(TOTAL_AMT - DISC_AMT, 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    oApp.ActiveSheet.Name = "B2B"
    oApp.Columns("A:N").EntireColumn.AutoFit
        
    '===========
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
'    Set oWB = oApp.Workbooks.Add
'    ActiveSheet.Name = "NewSheet"
    
    'Set oWB = oApp.Workbooks.Add
    'Sheets.Add After:=Sheets(Sheets.COUNT)

    Set oWS = oWB.Worksheets.Add
    Set oWS = oWB.Worksheets(1)
    oWS.Name = "Service Bills"
    
    oWS.Range("A" & 1).Value = "Invoice No."
    oWS.Range("B" & 1).Value = "Customer Name"
    oWS.Range("C" & 1).Value = "GSTIN"
    oWS.Range("D" & 1).Value = "Invoice Date"
    oWS.Columns("D:D").NumberFormat = "dd/mm/yyy"
    oWS.Range("E" & 1).Value = "Invoice Value"
    oWS.Range("F" & 1).Value = "Tax Rate(%)"
    oWS.Range("G" & 1).Value = "Taxable value"
    oWS.Range("H" & 1).Value = "IGST"
    oWS.Range("I" & 1).Value = "Central Tax"
    oWS.Range("J" & 1).Value = "State Tax"
    oWS.Range("K" & 1).Value = "Cess"
    oWS.Range("L" & 1).Value = "State of supply"
    oWS.Range("M" & 1).Value = "Reverse Charge"
    oWS.Range("N" & 1).Value = "E-com GSTIN"
    
    TOTAL_AMT = 0
    DISC_AMT = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    M = 1
    Do Until rstTRANX.EOF
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            M = M + 1
            oWS.Range("A" & M).Value = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                oWS.Range("B" & M).Value = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            Else
                oWS.Range("B" & M).Value = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            End If
            oWS.Range("C" & M).Value = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            oWS.Range("D" & M).Value = IIf(IsDate(rstTRANX!VCH_DATE), Format(rstTRANX!VCH_DATE, "DD/MM/YYYY"), "")
            CESSPER = 0
            CESSAMT = 0
            TaxAmt = 0
            TOTAL_AMT = 0
            TAXSALEAMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTtax!SALES_TAX & "  AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) - ((RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100)
                End Select
                If IsNull(RSTTRXFILE!QTY) Or RSTTRXFILE!QTY = 0 Then
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100)
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt)
                Else
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag) * Val(RSTTRXFILE!QTY)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100) * RSTTRXFILE!QTY
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * RSTTRXFILE!QTY * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt) * RSTTRXFILE!QTY
                End If
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            oWS.Range("F" & M).Value = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
            oWS.Range("G" & M).Value = Round(TAXSALEAMT, 2)
            If rstTRANX!CUST_IGST = "Y" Or (Len(Trim(oWS.Range("C" & M).Value)) = 15 And Left(Trim(oWS.Range("C" & M).Value), 2) <> Trim(MDIMAIN.LBLSTATE.Caption)) Then
                oWS.Range("H" & M).Value = Round(TaxAmt, 2)
            Else
                oWS.Range("I" & M).Value = Round(TaxAmt / 2, 2)
                oWS.Range("J" & M).Value = Round(TaxAmt / 2, 2)
            End If
            oWS.Range("K" & M).Value = Round(CESSPER + CESSAMT, 2)
            DISC_AMT = 0
            'If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            'End If
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                TOTAL_AMT = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
            oWS.Range("E" & M).Value = Round(TOTAL_AMT - DISC_AMT, 2)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    'Set oWS = Sheets.Add(After:=Sheets(Sheets.COUNT))
    oWS.Name = "SERVICE BILLS"
    oWS.Columns("A:Z").EntireColumn.AutoFit
    
    '=====
    '======
    Dim FIRST_BILL As Double
    Dim LAST_BILL As Double
    Dim FROMDATE As Date
    Dim TODATE As Date
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    GRDTranx.rows = 1
    GRDTranx.Cols = 9 + rstTRANX.RecordCount * 2
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = ""
    GRDTranx.TextMatrix(0, 2) = ""
    GRDTranx.TextMatrix(0, 3) = "BILL NOS"
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 0 '3500
    GRDTranx.ColWidth(2) = 0 '1800
    GRDTranx.ColWidth(3) = 3500
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 0 '3500
    GrdTotal.ColWidth(2) = 0 '1800
    GrdTotal.ColWidth(3) = 3500
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1300
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    GrdTotal.ColAlignment(5) = 4
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If GRDTranx.rows = 9 Then GoTo SKIPB2C
    n = 6
    M = 1
            
    
    Dim KFC As Double
    FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
    Do Until FROMDATE > TODATE
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            TOTAL_AMT = 0
            CESSPER = 0
            CESSAMT = 0
            KFC = 0
            rstTRANX.MoveLast
            LAST_BILL = rstTRANX!VCH_NO
            rstTRANX.MoveFirst
            FIRST_BILL = rstTRANX!VCH_NO
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = "" 'IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            GRDTranx.TextMatrix(M, 2) = "" 'IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(FIRST_BILL, "0000") & BILL_SUF & " TO " & BIL_PRE & Format(LAST_BILL, "0000") & BILL_SUF 'BIL_PRE & FIRST_BILL & BILL_SUF & " TO " & BIL_PRE & LAST_BILL & BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Do Until n = GRDTranx.Cols - 3
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)\
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                            Select Case RSTTRXFILE!SLSM_CODE
'                                Case "P"
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
'                                Case Else
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
'                            End Select
'                            GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        
                        Select Case RSTTRXFILE!SLSM_CODE
                            Case "P"
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                            Case Else
                                GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                        End Select
                        If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                            KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                            GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                            CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                            CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                        Else
                            KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                            GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                            CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                            CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                        End If
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                    Else
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                    End If
                    'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
            
            DISC_AMT = 0
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                If RSTtax!SLSM_CODE = "A" Then
                    DISC_AMT = DISC_AMT + IIf(IsNull(RSTtax!DISCOUNT), 0, RSTtax!DISCOUNT)
                End If
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        FROMDATE = DateAdd("d", FROMDATE, 1)
    Loop
    For i = 6 To GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    GrdTotal.rows = 0
    GrdTotal.rows = GrdTotal.rows + 1
    GrdTotal.Cols = GRDTranx.Cols
    For n = 4 To GRDTranx.Cols - 1
        If n <> 5 Then
            For i = 1 To GRDTranx.rows - 1
                GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
            Next i
        End If
    Next n
SKIPB2C:
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Cess Amount"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Addl Compensation Cess"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "KFC"
    'GrdTotal.FixedRows = 0
    'GrdTotal.Rows = 1
    
    '======
    'Set oWB = oApp.Workbooks.Add
    
    
    Set oWS = oWB.Worksheets.Add
    Set oWS = oWB.Worksheets(1)
    'Set oWS = Sheets.Add(After:=Sheets(Sheets.COUNT))
    oWS.Name = "B2C"
    'If Sum_flag = False Then
        oWS.Range("A1", "J1").Merge
        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
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
    
    oWS.Range("A" & 1).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (B2C SALES)"
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 2).Value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 2).Value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 2).Value = GRDTranx.TextMatrix(0, 2)
    oWS.Range("D" & 2).Value = GRDTranx.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 2).Value = GRDTranx.TextMatrix(0, 4)
    oWS.Range("F" & 2).Value = GRDTranx.TextMatrix(0, 5)
    oWS.Range("G" & 2).Value = GRDTranx.TextMatrix(0, 6)
    oWS.Range("H" & 2).Value = GRDTranx.TextMatrix(0, 7)
    oWS.Range("I" & 2).Value = GRDTranx.TextMatrix(0, 8)
    oWS.Range("J" & 2).Value = GRDTranx.TextMatrix(0, 9)
    oWS.Range("K" & 2).Value = GRDTranx.TextMatrix(0, 10)
    oWS.Range("L" & 2).Value = GRDTranx.TextMatrix(0, 11)
    oWS.Range("M" & 2).Value = GRDTranx.TextMatrix(0, 12)
    oWS.Range("N" & 2).Value = GRDTranx.TextMatrix(0, 13)
    oWS.Range("O" & 2).Value = GRDTranx.TextMatrix(0, 14)
    oWS.Range("P" & 2).Value = GRDTranx.TextMatrix(0, 15)
    oWS.Range("Q" & 2).Value = GRDTranx.TextMatrix(0, 16)
    oWS.Range("R" & 2).Value = GRDTranx.TextMatrix(0, 17)
    oWS.Range("S" & 2).Value = GRDTranx.TextMatrix(0, 18)
    oWS.Range("T" & 2).Value = GRDTranx.TextMatrix(0, 19)
    oWS.Range("U" & 2).Value = GRDTranx.TextMatrix(0, 20)
    oWS.Range("V" & 2).Value = GRDTranx.TextMatrix(0, 21)
    oWS.Range("W" & 2).Value = GRDTranx.TextMatrix(0, 22)
    oWS.Range("X" & 2).Value = GRDTranx.TextMatrix(0, 23)
    oWS.Range("Y" & 2).Value = GRDTranx.TextMatrix(0, 24)
    oWS.Range("Z" & 2).Value = GRDTranx.TextMatrix(0, 25)
    On Error GoTo ERRHAND
    
    i = 3
    For n = 1 To GRDTranx.rows - 1
        oWS.Range("A" & i).Value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDTranx.TextMatrix(n, 2)
        If IsDate(GRDTranx.TextMatrix(n, 3)) Then
            oWS.Range("D" & i).NumberFormat = "dd-mmm-yy"
            oWS.Range("D" & i).Value = Format(GRDTranx.TextMatrix(n, 3), "MM/dd/YYYY")
        Else
            oWS.Range("D" & i).Value = GRDTranx.TextMatrix(n, 3)
        End If
        'oWS.Range("D" & i).value = IIf(IsDate(GRDTranx.TextMatrix(N, 3)), Format(GRDTranx.TextMatrix(N, 3), "MM/dd/YYYY"), GRDTranx.TextMatrix(N, 3))
        On Error Resume Next
        oWS.Range("E" & i).Value = GRDTranx.TextMatrix(n, 4)
        If IsDate(GRDTranx.TextMatrix(n, 5)) Then
            oWS.Range("F" & i).NumberFormat = "dd-mmm-yy"
            oWS.Range("F" & i).Value = Format(GRDTranx.TextMatrix(n, 5), "MM/dd/YYYY")
        Else
            oWS.Range("F" & i).Value = GRDTranx.TextMatrix(n, 5)
        End If
        'oWS.Range("F" & i).value = IIf(IsDate(GRDTranx.TextMatrix(N, 5)), Format(GRDTranx.TextMatrix(N, 5), "MM/dd/YYYY"), GRDTranx.TextMatrix(N, 5))
        oWS.Range("G" & i).Value = GRDTranx.TextMatrix(n, 6)
        oWS.Range("H" & i).Value = GRDTranx.TextMatrix(n, 7)
        oWS.Range("I" & i).Value = GRDTranx.TextMatrix(n, 8)
        oWS.Range("J" & i).Value = GRDTranx.TextMatrix(n, 9)
        oWS.Range("K" & i).Value = GRDTranx.TextMatrix(n, 10)
        oWS.Range("L" & i).Value = GRDTranx.TextMatrix(n, 11)
        oWS.Range("M" & i).Value = GRDTranx.TextMatrix(n, 12)
        oWS.Range("N" & i).Value = GRDTranx.TextMatrix(n, 13)
        oWS.Range("O" & i).Value = GRDTranx.TextMatrix(n, 14)
        oWS.Range("P" & i).Value = GRDTranx.TextMatrix(n, 15)
        oWS.Range("Q" & i).Value = GRDTranx.TextMatrix(n, 16)
        oWS.Range("R" & i).Value = GRDTranx.TextMatrix(n, 17)
        oWS.Range("S" & i).Value = GRDTranx.TextMatrix(n, 18)
        oWS.Range("T" & i).Value = GRDTranx.TextMatrix(n, 19)
        oWS.Range("U" & i).Value = GRDTranx.TextMatrix(n, 20)
        oWS.Range("V" & i).Value = GRDTranx.TextMatrix(n, 21)
        oWS.Range("W" & i).Value = GRDTranx.TextMatrix(n, 22)
        oWS.Range("X" & i).Value = GRDTranx.TextMatrix(n, 23)
        oWS.Range("Y" & i).Value = GRDTranx.TextMatrix(n, 24)
        oWS.Range("Z" & i).Value = GRDTranx.TextMatrix(n, 25)
        On Error GoTo ERRHAND
        i = i + 1
    Next n
'    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
'    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    oApp.Selection.HorizontalAlignment = xlRight
'    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
'    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
'    oApp.Selection.Font.Bold = True
    
    If Sum_flag = True Then GoTo SKIPBTOC
'    oWS.Range("A" & i, "Z" & i).HorizontalAlignment = xlCenter
    oWS.Range("B" & i).Value = "TOTAL:"
    Dim k As Integer
    k = GRDTranx.Cols - 5
'    oWS.Range("C" & i).Formula = "=SUM(C4:C" & i - 1 & ")"
'    K = K - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("E" & i).Formula = "=SUM(E3:E" & i - 1 & ")"
'    K = K - 1
'    If K <= 0 Then GoTo SKIPBTOC
'    oWS.Range("F" & i).Formula = "=SUM(F3:F" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("G" & i).Formula = "=SUM(G3:G" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("H" & i).Formula = "=SUM(H3:H" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("I" & i).Formula = "=SUM(I3:I" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("J" & i).Formula = "=SUM(J3:J" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("K" & i).Formula = "=SUM(K3:K" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("L" & i).Formula = "=SUM(L3:L" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("M" & i).Formula = "=SUM(M3:M" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("N" & i).Formula = "=SUM(N3:N" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("O" & i).Formula = "=SUM(O3:O" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("P" & i).Formula = "=SUM(P3:P" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("Q" & i).Formula = "=SUM(Q3:Q" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("R" & i).Formula = "=SUM(R3:R" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("S" & i).Formula = "=SUM(S3:S" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("T" & i).Formula = "=SUM(T3:T" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("U" & i).Formula = "=SUM(U3:U" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("V" & i).Formula = "=SUM(V3:V" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("W" & i).Formula = "=SUM(W3:W" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("X" & i).Formula = "=SUM(X3:X" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("Y" & i).Formula = "=SUM(Y3:Y" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIPBTOC
    oWS.Range("Z" & i).Formula = "=SUM(Z3:Z" & i - 1 & ")"
    
    'oWS.Range("D" & i + 1).FormulaR1C1 = "=SUM(RC-10:RC-1)"
SKIPBTOC:
    'oApp.ActiveSheet.Name = "B2C"
    oWS.Columns("A:Z").EntireColumn.AutoFit
    
    '''=============
    ''PURCHASE REGISTER
    
    '''=============
    'SALES RETURN (CREDIT NOTE)
    
    
    
    
    '''=============
    'PURCHASE RETURN (DEBIT NOTE)
    
    '''==============
    'HSN WISE SALES
    
    '''==========
    oApp.Visible = True
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMDGSTR1_Click()
    Sum_flag = True
    GRDTranx.rows = 1
    GRDTranx.Cols = 14
    GrdTotal.rows = 0
    'GrdTotal.Cols = 14
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "Invoice No."
    GRDTranx.TextMatrix(0, 1) = "Customer Name"
    GRDTranx.TextMatrix(0, 2) = "GSTIN"
    GRDTranx.TextMatrix(0, 3) = "Invoice Date"
    GRDTranx.TextMatrix(0, 4) = "Invoice Value"
    GRDTranx.TextMatrix(0, 5) = "Tax Rate(%)"
    GRDTranx.TextMatrix(0, 6) = "Taxable value"
    GRDTranx.TextMatrix(0, 7) = "IGST"
    GRDTranx.TextMatrix(0, 8) = "Central Tax"
    GRDTranx.TextMatrix(0, 9) = "State Tax"
    GRDTranx.TextMatrix(0, 10) = "Cess"
    GRDTranx.TextMatrix(0, 11) = "State of supply"
    GRDTranx.TextMatrix(0, 12) = "Reverse Charge"
    GRDTranx.TextMatrix(0, 13) = "E-com GSTIN"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 1500
    GRDTranx.ColWidth(2) = 2500
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1000
    GRDTranx.ColWidth(6) = 1400
    GRDTranx.ColWidth(7) = 1200
    GRDTranx.ColWidth(8) = 1200
    GRDTranx.ColWidth(9) = 1200
    GRDTranx.ColWidth(10) = 1200
    GRDTranx.ColWidth(11) = 1700
    GRDTranx.ColWidth(12) = 1100
    GRDTranx.ColWidth(13) = 1100
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 1500
    GrdTotal.ColWidth(2) = 2500
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1400
    GrdTotal.ColWidth(5) = 1000
    GrdTotal.ColWidth(6) = 1400
    GrdTotal.ColWidth(7) = 1200
    GrdTotal.ColWidth(8) = 1200
    GrdTotal.ColWidth(9) = 1200
    GrdTotal.ColWidth(10) = 1200
    GrdTotal.ColWidth(11) = 1700
    GrdTotal.ColWidth(12) = 1100
    GrdTotal.ColWidth(13) = 1100
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    GRDTranx.ColAlignment(9) = 4
    GRDTranx.ColAlignment(10) = 4
    GRDTranx.ColAlignment(11) = 4
    GRDTranx.ColAlignment(12) = 4
    GRDTranx.ColAlignment(13) = 4
    
    GrdTotal.ColAlignment(0) = 4
    GrdTotal.ColAlignment(1) = 1
    GrdTotal.ColAlignment(2) = 1
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    GrdTotal.ColAlignment(5) = 4
    GrdTotal.ColAlignment(6) = 4
    GrdTotal.ColAlignment(7) = 4
    GrdTotal.ColAlignment(8) = 4
    GrdTotal.ColAlignment(9) = 4
    GrdTotal.ColAlignment(10) = 4
    GrdTotal.ColAlignment(11) = 4
    GrdTotal.ColAlignment(12) = 4
    GrdTotal.ColAlignment(13) = 4
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
    Dim CESSPER As Double
    Dim CESSAMT As Double
    Dim TOTAL_AMT As Double
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If OPTGST.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        ElseIf OptGR.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Optservice.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        Else
            BIL_PRE = ""
            BILL_SUF = ""
        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    If OPTGST.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    ElseIf OptGR.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    ElseIf Optservice.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    ElseIf OptRT.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    ElseIf Opt8V.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    ElseIf OptWs.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    M = 0
    Do Until rstTRANX.EOF
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            M = M + 1
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            
            CESSPER = 0
            CESSAMT = 0
            TaxAmt = 0
            TOTAL_AMT = 0
            TAXSALEAMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTtax!SALES_TAX & "  AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) - ((RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100)
                End Select
                
                If IsNull(RSTTRXFILE!QTY) Or RSTTRXFILE!QTY = 0 Then
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100)
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt)
                Else
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag) * Val(RSTTRXFILE!QTY)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100) * RSTTRXFILE!QTY
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * RSTTRXFILE!QTY * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt) * RSTTRXFILE!QTY
                End If
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            GRDTranx.TextMatrix(M, 5) = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
            GRDTranx.TextMatrix(M, 6) = TAXSALEAMT
            If rstTRANX!CUST_IGST = "Y" Or (Len(Trim(GRDTranx.TextMatrix(M, 2))) = 15 And Left(Trim(GRDTranx.TextMatrix(M, 2)), 2) <> Trim(MDIMAIN.LBLSTATE.Caption)) Then
                GRDTranx.TextMatrix(M, 7) = TaxAmt
            Else
                GRDTranx.TextMatrix(M, 8) = TaxAmt / 2
                GRDTranx.TextMatrix(M, 9) = TaxAmt / 2
            End If
            GRDTranx.TextMatrix(M, 10) = CESSPER + CESSAMT
            DISC_AMT = 0
            'If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            'End If
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                TOTAL_AMT = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
            GRDTranx.TextMatrix(M, 4) = TOTAL_AMT - DISC_AMT
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
        
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdJSON_Click()
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim Num As Currency
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "JSON\e_way.json" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "JSON\gstr.json" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    
    Dim CompTin As String
    Dim gross_amt As Double
    Dim DISC_AMT As Double
    Dim TAX_AMT As Double
    Dim tot_val As Double
    Dim cess_amt As Double
    Dim BIL_PRE, BILL_SUF, BIL_PRE_R, BILL_SUF_R As String
    
    Screen.MousePointer = vbHourglass
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", RSTCOMPANY!CST)
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        BIL_PRE_R = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        BILL_SUF_R = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim CESSPER As Double
    Dim CESSAMT As Double
    Dim TaxAmt As Double
    Dim TOTAL_AMT As Double
    Dim TAXSALEAMT As Double

    Dim M As Double
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    '{"ctin":"32AXMPR3460Q1ZT","inv":[{"inum":"960","idt":"01-02-2023","val":1010.00,"pos":"32","rchrg":"N","itms":[{"num":1,"itm_det":{"txval":961.96,"rt":5,"camt":24.05,"samt":24.05,"csamt":0.00}}],"inv_typ":"R"}]}
    
    Print #1, "{" & Chr(34) & "gstin" & Chr(34) & ":" & Chr(34) & CompTin & Chr(34) & "," & Chr(34) & "fp" & Chr(34) & ":" & Chr(34) & Format(Month(DTFROM.Value), "00") & Format(Year(DTFROM.Value), "0000") & Chr(34) & _
    "," & Chr(34); "gt" & Chr(34) & ":0.00" & "," & Chr(34) & "cur_gt" & Chr(34) & ":0.00" & "," & Chr(34) & "b2b" & Chr(34) & ":" & "["
    
    gross_amt = 0
    tot_val = 0
    DISC_AMT = 0
    cess_amt = 0
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI'  ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    M = 0
    Do Until rstTRANX.EOF
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            
            M = M + 1
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            
            CESSPER = 0
            CESSAMT = 0
            TaxAmt = 0
            TOTAL_AMT = 0
            TAXSALEAMT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTtax!SALES_TAX & "  AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) - ((RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100)
                End Select
                If IsNull(RSTTRXFILE!QTY) Or RSTTRXFILE!QTY = 0 Then
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100)
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt)
                Else
                    TAXSALEAMT = TAXSALEAMT + Val(GRDTranx.Tag) * Val(RSTTRXFILE!QTY)
                    TaxAmt = TaxAmt + (Val(GRDTranx.Tag) * RSTTRXFILE!SALES_TAX / 100) * RSTTRXFILE!QTY
                    CESSPER = CESSPER + (RSTTRXFILE!PTR - (RSTTRXFILE!PTR * RSTTRXFILE!LINE_DISC) / 100) * RSTTRXFILE!QTY * IIf(IsNull(RSTTRXFILE!CESS_PER), 0, RSTTRXFILE!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTTRXFILE!cess_amt), 0, RSTTRXFILE!cess_amt) * RSTTRXFILE!QTY
                End If
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            GRDTranx.TextMatrix(M, 5) = IIf(IsNull(RSTtax!SALES_TAX), "", RSTtax!SALES_TAX)
            GRDTranx.TextMatrix(M, 6) = TAXSALEAMT
            If rstTRANX!CUST_IGST = "Y" Or (Len(Trim(GRDTranx.TextMatrix(M, 2))) = 15 And Left(Trim(GRDTranx.TextMatrix(M, 2)), 2) <> Trim(MDIMAIN.LBLSTATE.Caption)) Then
                GRDTranx.TextMatrix(M, 7) = TaxAmt
            Else
                GRDTranx.TextMatrix(M, 8) = TaxAmt / 2
                GRDTranx.TextMatrix(M, 9) = TaxAmt / 2
            End If
            GRDTranx.TextMatrix(M, 10) = CESSPER + CESSAMT
            DISC_AMT = 0
            'If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            'End If
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                TOTAL_AMT = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
            
            GRDTranx.TextMatrix(M, 4) = TOTAL_AMT - DISC_AMT
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
        
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
    
    Print #1, "]"
    Print #1, "}"
    Close #1 '//Closing the file
    
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CMDREGISTER_Click()

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
    If Optsales.Value = True Then
        If OPTGST.Value = True Then
            oWS.Range("A" & 2).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (" & OPTGST.Caption & ")"
        ElseIf OptGR.Value = True Then
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                oWS.Range("A" & 2).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
            Else
                oWS.Range("A" & 2).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (" & OptGR.Caption & ")"
            End If
        ElseIf Optservice.Value = True Then
            oWS.Range("A" & 2).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (SERVICE BILLS)"
        ElseIf Optst.Value = True Then
            oWS.Range("A" & 2).Value = "REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (Delivery Chellans)"
        Else
            oWS.Range("A" & 2).Value = "SALES REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " (OTHERS)"
        End If
    ElseIf OptPurchase.Value = True Then
        If OptNormal.Value = True Then
            oWS.Range("A" & 2).Value = "PURCHASE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
        Else
            oWS.Range("A" & 2).Value = "LOCAL PURCHASE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
        End If
    ElseIf OptCST.Value = True Then
        oWS.Range("A" & 2).Value = "CST BILL REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    ElseIf OptPurchret.Value = True Then
        oWS.Range("A" & 2).Value = "DEBIT NOTE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    ElseIf OptExpense.Value = True Then
        oWS.Range("A" & 2).Value = "EXPENSE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    ElseIf OptExReturn.Value = True Then
        oWS.Range("A" & 2).Value = "EXCHANGE GOODS REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    ElseIf optAssets.Value = True Then
        oWS.Range("A" & 2).Value = "ASSETS PURCHASE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    ElseIf OptDamage.Value = True Then
        oWS.Range("A" & 2).Value = "DAMAGE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    Else
        oWS.Range("A" & 2).Value = "CREDIT NOTE REGISTER FOR THE PERIOD FROM " & DTFROM.Value & " TO " & DTTo.Value & " "
    End If
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDTranx.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GRDTranx.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GRDTranx.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GRDTranx.TextMatrix(0, 5)
    oWS.Range("G" & 3).Value = GRDTranx.TextMatrix(0, 6)
    oWS.Range("H" & 3).Value = GRDTranx.TextMatrix(0, 7)
    oWS.Range("I" & 3).Value = GRDTranx.TextMatrix(0, 8)
    oWS.Range("J" & 3).Value = GRDTranx.TextMatrix(0, 9)
    oWS.Range("K" & 3).Value = GRDTranx.TextMatrix(0, 10)
    oWS.Range("L" & 3).Value = GRDTranx.TextMatrix(0, 11)
    oWS.Range("M" & 3).Value = GRDTranx.TextMatrix(0, 12)
    oWS.Range("N" & 3).Value = GRDTranx.TextMatrix(0, 13)
    oWS.Range("O" & 3).Value = GRDTranx.TextMatrix(0, 14)
    oWS.Range("P" & 3).Value = GRDTranx.TextMatrix(0, 15)
    oWS.Range("Q" & 3).Value = GRDTranx.TextMatrix(0, 16)
    oWS.Range("R" & 3).Value = GRDTranx.TextMatrix(0, 17)
    oWS.Range("S" & 3).Value = GRDTranx.TextMatrix(0, 18)
    oWS.Range("T" & 3).Value = GRDTranx.TextMatrix(0, 19)
    oWS.Range("U" & 3).Value = GRDTranx.TextMatrix(0, 20)
    oWS.Range("V" & 3).Value = GRDTranx.TextMatrix(0, 21)
    oWS.Range("W" & 3).Value = GRDTranx.TextMatrix(0, 22)
    oWS.Range("X" & 3).Value = GRDTranx.TextMatrix(0, 23)
    oWS.Range("Y" & 3).Value = GRDTranx.TextMatrix(0, 24)
    oWS.Range("Z" & 3).Value = GRDTranx.TextMatrix(0, 25)
    On Error GoTo ERRHAND
    
    i = 4
    For n = 1 To GRDTranx.rows - 1
        oWS.Range("A" & i).NumberFormat = "@"
        oWS.Range("A" & i).Value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDTranx.TextMatrix(n, 2)
        
        'oWS.Range("D" & i).Value = GRDTranx.TextMatrix(n, 3)
'        If IsDate(GRDTranx.TextMatrix(n, 3)) And Val(GRDTranx.TextMatrix(n, 3)) <> 0 And Len(GRDTranx.TextMatrix(n, 3)) = 10 Then
'            oWS.Range("D" & i).value = Format(GRDTranx.TextMatrix(n, 3), "MM/dd/YYYY")
'        Else
'            oWS.Range("D" & i).value = GRDTranx.TextMatrix(n, 3)
'        End If
        'oWS.Range("D" & i).value = IIf(IsDate(GRDTranx.TextMatrix(N, 3)), Format(GRDTranx.TextMatrix(N, 3), "MM/dd/YYYY"), GRDTranx.TextMatrix(N, 3))
        On Error Resume Next
        If IsDate(GRDTranx.TextMatrix(n, 3)) And Val(GRDTranx.TextMatrix(n, 3)) <> 0 And Len(GRDTranx.TextMatrix(n, 3)) = 10 Then
            oWS.Range("D" & i).NumberFormat = "dd-mmm-yy"
            oWS.Range("D" & i).Value = Format(GRDTranx.TextMatrix(n, 3), "MM/dd/YYYY")
        Else
            oWS.Range("D" & i).NumberFormat = "@"
            oWS.Range("D" & i).Value = GRDTranx.TextMatrix(n, 3)
        End If
        oWS.Range("E" & i).Value = GRDTranx.TextMatrix(n, 4)
        If IsDate(GRDTranx.TextMatrix(n, 5)) And Val(GRDTranx.TextMatrix(n, 5)) <> 0 And Len(GRDTranx.TextMatrix(n, 5)) = 10 Then
            oWS.Range("F" & i).NumberFormat = "dd-mmm-yy"
            oWS.Range("F" & i).Value = Format(GRDTranx.TextMatrix(n, 5), "MM/dd/YYYY")
        Else
            oWS.Range("F" & i).Value = GRDTranx.TextMatrix(n, 5)
        End If
        'If IsDate(GRDTranx.TextMatrix(N, 6)) Then
        If IsDate(GRDTranx.TextMatrix(n, 6)) And Val(GRDTranx.TextMatrix(n, 6)) <> 0 And Len(GRDTranx.TextMatrix(n, 6)) = 10 Then
            oWS.Range("G" & i).NumberFormat = "dd-mmm-yy"
            oWS.Range("G" & i).Value = Format(GRDTranx.TextMatrix(n, 6), "MM/dd/YYYY")
        Else
            oWS.Range("G" & i).Value = GRDTranx.TextMatrix(n, 6)
        End If
        
        'oWS.Range("F" & i).value = IIf(IsDate(GRDTranx.TextMatrix(N, 5)), Format(GRDTranx.TextMatrix(N, 5), "MM/dd/YYYY"), GRDTranx.TextMatrix(N, 5))
        'oWS.Range("G" & i).value = GRDTranx.TextMatrix(N, 6)
        oWS.Range("H" & i).Value = GRDTranx.TextMatrix(n, 7)
        oWS.Range("I" & i).Value = GRDTranx.TextMatrix(n, 8)
        oWS.Range("J" & i).Value = GRDTranx.TextMatrix(n, 9)
        oWS.Range("K" & i).Value = GRDTranx.TextMatrix(n, 10)
        oWS.Range("L" & i).Value = GRDTranx.TextMatrix(n, 11)
        oWS.Range("M" & i).Value = GRDTranx.TextMatrix(n, 12)
        oWS.Range("N" & i).Value = GRDTranx.TextMatrix(n, 13)
        oWS.Range("O" & i).Value = GRDTranx.TextMatrix(n, 14)
        oWS.Range("P" & i).Value = GRDTranx.TextMatrix(n, 15)
        oWS.Range("Q" & i).Value = GRDTranx.TextMatrix(n, 16)
        oWS.Range("R" & i).Value = GRDTranx.TextMatrix(n, 17)
        oWS.Range("S" & i).Value = GRDTranx.TextMatrix(n, 18)
        oWS.Range("T" & i).Value = GRDTranx.TextMatrix(n, 19)
        oWS.Range("U" & i).Value = GRDTranx.TextMatrix(n, 20)
        oWS.Range("V" & i).Value = GRDTranx.TextMatrix(n, 21)
        oWS.Range("W" & i).Value = GRDTranx.TextMatrix(n, 22)
        oWS.Range("X" & i).Value = GRDTranx.TextMatrix(n, 23)
        oWS.Range("Y" & i).Value = GRDTranx.TextMatrix(n, 24)
        oWS.Range("Z" & i).Value = GRDTranx.TextMatrix(n, 25)
        On Error GoTo ERRHAND
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
    If Sum_flag = True Then GoTo SKIP
'    oWS.Range("A" & i, "Z" & i).HorizontalAlignment = xlCenter
    oWS.Range("B" & i).Value = "TOTAL:"
    Dim k As Integer
    k = GRDTranx.Cols - 4
'    oWS.Range("C" & i).Formula = "=SUM(C4:C" & i - 1 & ")"
'    K = K - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("E" & i).Formula = "=SUM(E4:E" & i - 1 & ")"
'    K = K - 1
'    If K <= 0 Then GoTo SKIP
'    oWS.Range("F" & i).Formula = "=SUM(F4:F" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("G" & i).Formula = "=SUM(G4:G" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("H" & i).Formula = "=SUM(H4:H" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("I" & i).Formula = "=SUM(I4:I" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("J" & i).Formula = "=SUM(J4:J" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("K" & i).Formula = "=SUM(K4:K" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("L" & i).Formula = "=SUM(L4:L" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("M" & i).Formula = "=SUM(M4:M" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("N" & i).Formula = "=SUM(N4:N" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("O" & i).Formula = "=SUM(O4:O" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("P" & i).Formula = "=SUM(P4:P" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("Q" & i).Formula = "=SUM(Q4:Q" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("R" & i).Formula = "=SUM(R4:R" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("S" & i).Formula = "=SUM(S4:S" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("T" & i).Formula = "=SUM(T4:T" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("U" & i).Formula = "=SUM(U4:U" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("V" & i).Formula = "=SUM(V4:V" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("W" & i).Formula = "=SUM(W4:W" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("X" & i).Formula = "=SUM(X4:X" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("Y" & i).Formula = "=SUM(Y4:Y" & i - 1 & ")"
    k = k - 1
    If k <= 0 Then GoTo SKIP
    oWS.Range("Z" & i).Formula = "=SUM(Z4:Z" & i - 1 & ")"
    
    'oWS.Range("D" & i + 1).FormulaR1C1 = "=SUM(RC-10:RC-1)"
SKIP:
    oApp.Visible = True
    
    If Sum_flag = True Then
        'oWS.Columns("C:C").Select
        oWS.Columns("C:C").NumberFormat = "0"
        oWS.Columns("A:Z").EntireColumn.AutoFit
    End If
    
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

Private Sub CmdReset_Click()
    If MsgBox("Are you sure you want to Reset the Rates", vbYesNo, "B2C Sales") = vbNo Then Exit Sub
    db.Execute "UPDATE TRXFILE SET ST_RATE = 0 WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' "
End Sub

Private Sub CmdSummary_Click()
    Sum_flag = False
    If Optsales.Value = False Then Exit Sub
    Dim i, n As Long
    On Error GoTo ERRHAND
    If Optbrsale.Value = True Then
        Call Sales_Register_DailyBR
    Else
        Call Sales_Register_Daily
    End If
'    GrdTotal.Rows = 0
'    GrdTotal.Rows = GrdTotal.Rows + 1
'    GrdTotal.Cols = GRDTranx.Cols
'    GrdTotal.TextMatrix(0, 3) = "TOTAL"
'    For N = 4 To GRDTranx.Cols - 1
'        If N <> 5 Then
'            For i = 1 To GRDTranx.Rows - 1
'                GrdTotal.TextMatrix(0, N) = Val(GrdTotal.TextMatrix(0, N)) + Val(GRDTranx.TextMatrix(i, N))
'            Next i
'        End If
'    Next N
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
        
    db.Execute "Delete from hsn_trxfile "
    db.Execute "Delete from hsn_trxmast "
    
    db.Execute "INSERT INTO `hsn_trxfile` SELECT * FROM `trxfile` WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV' OR TRX_TYPE='DM' OR TRX_TYPE='DG') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y')"
    db.Execute "INSERT INTO `hsn_trxmast` SELECT * FROM `trxmast` WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') "
    
    Dim rstTRANX As ADODB.Recordset
    Dim rsthsnTRANX As ADODB.Recordset
    Dim rsthsnTRXFILE As ADODB.Recordset
    Dim rsthsnTRXMAST As ADODB.Recordset
    
    Set rsthsnTRXFILE = New ADODB.Recordset
    rsthsnTRXFILE.Open "Select * From DAMAGE_MAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='DM' OR TRX_TYPE='DG') ", db, adOpenStatic, adLockReadOnly
    
    Set rsthsnTRXMAST = New ADODB.Recordset
    rsthsnTRXMAST.Open "SELECT * From hsn_trxmast", db, adOpenStatic, adLockOptimistic, adCmdText
    rsthsnTRXMAST.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rsthsnTRXFILE.EOF
        rsthsnTRXMAST.AddNew
        rsthsnTRXMAST!TRX_TYPE = rsthsnTRXFILE!TRX_TYPE
        rsthsnTRXMAST!VCH_NO = rsthsnTRXFILE!VCH_NO
        rsthsnTRXMAST!TRX_YEAR = rsthsnTRXFILE!TRX_YEAR
        rsthsnTRXMAST!VCH_DATE = rsthsnTRXFILE!VCH_DATE
        rsthsnTRXMAST!DISC_PERS = 0
        rsthsnTRXMAST!CUST_IGST = ""
        rsthsnTRXMAST!TIN = IIf(IsNull(rsthsnTRXFILE!TIN), "", rsthsnTRXFILE!TIN)
        rsthsnTRXMAST.Update
        
        rsthsnTRXFILE.MoveNext
    Loop
    rsthsnTRXMAST.Close
    Set rsthsnTRXMAST = Nothing
        
    rsthsnTRXFILE.Close
    Set rsthsnTRXFILE = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    'rstTRANX.Open "SELECT * From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
    rstTRANX.Open "SELECT * From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SV') AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly
    Set rsthsnTRANX = New ADODB.Recordset
    rsthsnTRANX.Open "SELECT * From hsn_trxfile", db, adOpenStatic, adLockOptimistic, adCmdText
    rsthsnTRANX.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rstTRANX.EOF
        rsthsnTRANX.AddNew
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                rsthsnTRANX!TRX_TYPE = "E1"
            Case "HI"
                rsthsnTRANX!TRX_TYPE = "E2"
            Case "SV"
                rsthsnTRANX!TRX_TYPE = "E3"
            Case Else
                rsthsnTRANX!TRX_TYPE = rstTRANX!TRX_TYPE
        End Select
        rsthsnTRANX!VCH_NO = rstTRANX!VCH_NO
        rsthsnTRANX!TRX_YEAR = rstTRANX!TRX_YEAR
        rsthsnTRANX!VCH_DATE = rstTRANX!VCH_DATE
        rsthsnTRANX!LINE_NO = rstTRANX!LINE_NO
        rsthsnTRANX!ITEM_CODE = rstTRANX!ITEM_CODE
        rsthsnTRANX!ITEM_NAME = rstTRANX!ITEM_NAME
        If IsNull(rstTRANX!QTY) Or rstTRANX!QTY = 0 Then
            rsthsnTRANX!QTY = 0
            rsthsnTRANX!PTR = Round(((rstTRANX!TRX_TOTAL * 100) / ((rstTRANX!SALES_TAX) + 100)), 3) 'rstTRANX!PTR
        Else
            rsthsnTRANX!QTY = -rstTRANX!QTY
            rsthsnTRANX!PTR = Round(((rstTRANX!TRX_TOTAL * 100) / ((rstTRANX!SALES_TAX) + 100)) / rstTRANX!QTY, 3) 'rstTRANX!PTR
        End If
        rsthsnTRANX!SALES_TAX = rstTRANX!SALES_TAX
        rsthsnTRANX!TRX_TOTAL = rstTRANX!TRX_TOTAL
        rsthsnTRANX!LINE_DISC = 0 'rstTRANX!P_DISC
        rsthsnTRANX!FREE_QTY = rstTRANX!FREE_QTY
        rsthsnTRANX!LOOSE_PACK = rstTRANX!LOOSE_PACK
        rsthsnTRANX!cess_amt = IIf(IsNull(rstTRANX!cess_amt), 0, rstTRANX!cess_amt)
        rsthsnTRANX!CESS_PER = IIf(IsNull(rstTRANX!CESS_PER), 0, rstTRANX!CESS_PER)
        rsthsnTRANX.Update

        rstTRANX.MoveNext
    Loop
    rsthsnTRANX.Close
    Set rsthsnTRANX = Nothing

    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
    Set rsthsnTRXFILE = New ADODB.Recordset
    rsthsnTRXFILE.Open "Select * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ", db, adOpenStatic, adLockReadOnly
    
    Set rsthsnTRXMAST = New ADODB.Recordset
    rsthsnTRXMAST.Open "SELECT * From hsn_trxmast", db, adOpenStatic, adLockOptimistic, adCmdText
    rsthsnTRXMAST.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rsthsnTRXFILE.EOF
        rsthsnTRXMAST.AddNew
        rsthsnTRXMAST!TRX_TYPE = rsthsnTRXFILE!TRX_TYPE
        rsthsnTRXMAST!VCH_NO = rsthsnTRXFILE!VCH_NO
        rsthsnTRXMAST!TRX_YEAR = rsthsnTRXFILE!TRX_YEAR
        rsthsnTRXMAST!VCH_DATE = rsthsnTRXFILE!VCH_DATE
        rsthsnTRXMAST!DISC_PERS = 0
        rsthsnTRXMAST!CUST_IGST = ""
        rsthsnTRXMAST!TIN = IIf(IsNull(rsthsnTRXFILE!TIN), "", rsthsnTRXFILE!TIN)
        rsthsnTRXMAST.Update
        
        rsthsnTRXFILE.MoveNext
    Loop
    rsthsnTRXMAST.Close
    Set rsthsnTRXMAST = Nothing
        
    rsthsnTRXFILE.Close
    Set rsthsnTRXFILE = Nothing
    
    
    Set rsthsnTRXFILE = New ADODB.Recordset
    rsthsnTRXFILE.Open "SELECT DISTINCT TRX_TYPE, VCH_NO, TRX_YEAR From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
    
    Set rsthsnTRXMAST = New ADODB.Recordset
    rsthsnTRXMAST.Open "SELECT * From hsn_trxmast", db, adOpenStatic, adLockOptimistic, adCmdText
    rsthsnTRXMAST.Properties("Update Criteria").Value = adCriteriaKey
    Do Until rsthsnTRXFILE.EOF
        rsthsnTRXMAST.AddNew
        Select Case rsthsnTRXFILE!TRX_TYPE
            Case "GI"
                rsthsnTRXMAST!TRX_TYPE = "E1"
            Case "HI"
                rsthsnTRXMAST!TRX_TYPE = "E2"
            Case "SV"
                rsthsnTRXMAST!TRX_TYPE = "E3"
            Case Else
                rsthsnTRXMAST!TRX_TYPE = rsthsnTRXFILE!TRX_TYPE
        End Select
        rsthsnTRXMAST!VCH_NO = rsthsnTRXFILE!VCH_NO
        rsthsnTRXMAST!TRX_YEAR = rsthsnTRXFILE!TRX_YEAR
        'rsthsnTRXMAST!VCH_DATE = rsthsnTRXFILE!VCH_DATE
        rsthsnTRXMAST!DISC_PERS = 0
        rsthsnTRXMAST!CUST_IGST = ""
        rsthsnTRXMAST!TIN = ""
        rsthsnTRXMAST.Update
        
        rsthsnTRXFILE.MoveNext
    Loop
    rsthsnTRXMAST.Close
    Set rsthsnTRXMAST = Nothing
    
    If optcombine.Value = True Then
        ReportNameVar = Rptpath & "RPTHSNREPORT"
    Else
        ReportNameVar = Rptpath & "RPTHSNREPORTSP"
    End If
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV' OR {TRXFILE.TRX_TYPE}='SR' OR {TRXFILE.TRX_TYPE}='EX' OR {TRXFILE.TRX_TYPE}='E1' OR {TRXFILE.TRX_TYPE}='E2' OR {TRXFILE.TRX_TYPE}='E3' OR {TRXFILE.TRX_TYPE} ='DM' OR {TRXFILE.TRX_TYPE} ='DG' ) AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        
    
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command2_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    db.Execute "Update trxfile set CESS_PER =0 where isnull(CESS_PER)"
    db.Execute "Update trxfile set CESS_AMT =0 where isnull(CESS_AMT)"
    
    If optcombine.Value = True Then
        ReportNameVar = Rptpath & "RPTSALESREPORT"
    Else
        ReportNameVar = Rptpath & "RPTSALESREPORTSP"
    End If
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OptRT.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='RI') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptGR.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Optservice.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='SV')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Optst.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='TF')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Opt8V.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='VI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptWs.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OPTGST.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command3_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    ReportNameVar = Rptpath & "RPTPURCHASEREPORT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OptNormal.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='PI' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptComm.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='PW' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptLocal.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='LP' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command4_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
        
    'On Error GoTo ErrHand
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    
    If optcombine.Value = True Then
        ReportNameVar = Rptpath & "RPTITEMHSNREPORT"
    Else
        ReportNameVar = Rptpath & "RPTITEMHSNREPORTSP"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTGST.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptGR.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptRT.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='RI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Optservice.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='SV')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptWs.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='SI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf Optst.Value = True Then
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='TF')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='SV') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command5_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    ReportNameVar = Rptpath & "RPTITEMPURCHASE"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OptNormal.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='PI' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptComm.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='PW' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    ElseIf OptLocal.Value = True Then
        Report.RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='LP' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command6_Click()
    Call CMDGSTR1_Click
End Sub

Private Sub Command7_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    If optcombine.Value = True Then
        ReportNameVar = Rptpath & "RPTHSNREPORTSM"
    Else
        ReportNameVar = Rptpath & "RPTHSNREPORTSMSP"
    End If
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI' OR {TRXFILE.TRX_TYPE}='SV')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        
    
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "ITEM WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Command8_Click()
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    If OptComb1.Value = True Then
        ReportNameVar = Rptpath & "RPTHSNREPORTSP1"
    Else
        ReportNameVar = Rptpath & "RPTHSNREPORTSP2"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "((ISNULL({TRXFILE.UN_BILL}) OR {TRXFILE.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' OR {TRXFILE.TRX_TYPE}='HI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
    Next
    frmreport.Caption = "HSN WISE SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
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
    Frame1.Visible = False
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        Optservice.Caption = "Sales Bill(M)"
        OPTGST.Visible = False
        OptGR.Caption = "Sales Bill"
    End If
    Sum_flag = False
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 0
    Me.Top = 0
    ACT_FLAG = True
    PHY_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDTranx_DblClick()
    If frmLogin.rs!Level = "5" Then Exit Sub
    If Optsales.Value = False Then Exit Sub
    If OptGR.Value = True Then
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
            If IsFormLoaded(frmsales) <> True Then
                frmsales.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                frmsales.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                frmsales.Show
                frmsales.SetFocus
                Call frmsales.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMSALES1) <> True Then
                FRMSALES1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMSALES1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMSALES1.Show
                FRMSALES1.SetFocus
                Call FRMSALES1.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMSALES2) <> True Then
                FRMSALES2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMSALES2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                FRMSALES2.Show
                FRMSALES2.SetFocus
                Call FRMSALES2.txtBillNo_KeyDown(13, 0)
            End If
        Else
            If SALESLT_FLAG = "Y" Then
                If IsFormLoaded(FRMGSTRSM1) <> True Then
                    FRMGSTRSM1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM1.Show
                    FRMGSTRSM1.SetFocus
                    Call FRMGSTRSM1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGSTRSM2) <> True Then
                    FRMGSTRSM2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM2.Show
                    FRMGSTRSM2.SetFocus
                    Call FRMGSTRSM2.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGSTRSM3) <> True Then
                    FRMGSTRSM3.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM3.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTRSM3.Show
                    FRMGSTRSM3.SetFocus
                    Call FRMGSTRSM3.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If IsFormLoaded(FRMGSTR) <> True Then
                    FRMGSTR.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR.Show
                    FRMGSTR.SetFocus
                    Call FRMGSTR.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGSTR1) <> True Then
                    FRMGSTR1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR1.Show
                    FRMGSTR1.SetFocus
                    Call FRMGSTR1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGSTR2) <> True Then
                    FRMGSTR2.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR2.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGSTR2.Show
                    FRMGSTR2.SetFocus
                    Call FRMGSTR2.txtBillNo_KeyDown(13, 0)
                End If
            End If
        End If
    ElseIf OPTGST.Value = True Then
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDTranx.TextMatrix(GRDTranx.Row, 15)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                Exit Sub
            Else
                If IsFormLoaded(FRMGST) <> True Then
                    FRMGST.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST.Show
                    FRMGST.SetFocus
                    Call FRMGST.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMGST1) <> True Then
                    FRMGST1.txtBillNo.Text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST1.LBLBILLNO.Caption = Val(GRDTranx.TextMatrix(GRDTranx.Row, 3))
                    FRMGST1.Show
                    FRMGST1.SetFocus
                    Call FRMGST1.txtBillNo_KeyDown(13, 0)
                End If
            End If
    End If
End Sub

Private Sub LBLTOTAL_DblClick(index As Integer)
    If Frame1.Visible = True Then
        Frame1.Visible = False
    Else
        Frame1.Visible = True
    End If
End Sub

Private Sub optAssets_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Assets Register"
End Sub

Private Sub OptCST_Click()
    FrameSales.Visible = False
    Frmeperiod.Caption = "CST Bill Register"
End Sub

Private Sub OPTEXPENSE_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Expense Register"
End Sub

Private Sub OptExReturn_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Exchange Register"
End Sub

Private Sub optPurchase_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = True
    Frmeperiod.Caption = "Purchase Register"
End Sub

Private Sub OptPurchret_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Purchase Return Register"
End Sub

Private Sub Optsales_Click()
    FrameSales.Visible = True
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Sales Register"
End Sub

Private Function Sales_Register()
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If OPTGST.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        ElseIf OptGR.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Optservice.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        Else
            BIL_PRE = ""
            BILL_SUF = ""
        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optst.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='VI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf optcombine.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        End If
        GRDTranx.rows = 1
        GRDTranx.Cols = 10 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optst.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                Else
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Or rstTRANX!BILL_NAME = "CASH" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            'Trim(txtBillNo.text)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            Dim KFC As Double
            TOTAL_AMT = 0
            KFC = 0
            Do Until n = GRDTranx.Cols - 4
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                    End If
                    
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 3)
            GRDTranx.TextMatrix(M, n + 3) = Format(Round(DISC_AMT, 3), "0.00")
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 2
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "KFC"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "DISC"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 4) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optst.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL), 3)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing

            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Purchase_Register()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    'On Error GoTo ErrHand
    
    Set rstTRANX = New ADODB.Recordset
    If OptNormal.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    ElseIf OptComm.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    End If
    GRDTranx.rows = 1
    GRDTranx.Cols = (6 + rstTRANX.RecordCount * 2) + 4 + 1 + 1
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Supplier."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1500
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Ser. Taxable Amt"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Ser Tax Amt"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Cess"
    
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Ser. Taxable Amt"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Ser Tax Amt"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 6) = "Cess"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 5) = "Addl Cess"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Discount"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "TCS"
    
'    GrdTotal.ColWidth(GRDTranx.Cols - 2) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 1) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 4) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 3) = 1500
    
    GrdTotal.ColWidth(GRDTranx.Cols - 4) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 3) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 5) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 6) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 2) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 1) = 1500
    
    GrdTotal.ColAlignment(GRDTranx.Cols - 1) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 2) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 6) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 4) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 3) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 5) = 4
                    
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If GRDTranx.rows = 6 Then Exit Function
    
    n = 6
    M = 1
    Dim TAX_PER As Single
    Dim CESSPER As Double
    Dim CESSAMT As Double
    
    Dim RSTACTMAST As ADODB.Recordset
    Set rstTRANX = New ADODB.Recordset
    If OptNormal.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    ElseIf OptComm.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    End If
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!ACT_NAME
        
        Set RSTACTMAST = New ADODB.Recordset
        RSTACTMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
        End If
        RSTACTMAST.Close
        Set RSTACTMAST = Nothing
        
        GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!PINV), "", rstTRANX!PINV)
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), "", rstTRANX!NET_AMOUNT)
        If Optinvdate.Value = True Then
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        Else
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!RCVD_DATE), rstTRANX!RCVD_DATE, "")
        End If
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = IIf(IsNull(rstTRANX!ADD_AMOUNT), 0, rstTRANX!ADD_AMOUNT)
        CESSPER = 0
        CESSAMT = 0
        Do Until n = GRDTranx.Cols - 4
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & Val(GRDTranx.TextMatrix(0, n)) & " AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                
                Select Case RSTtax!DISC_FLAG
                    Case "P"
                        GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!TR_DISC / 100) '- ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY)) - ((RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY))) * RSTtax!TR_DISC / 100)
                End Select
                GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))


                TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
                If RSTtax!DISC_FLAG = "P" Then
                    CESSPER = CESSPER + (Val(GRDTranx.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                Else
                    CESSPER = CESSPER + (Val(GRDTranx.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                End If
                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
'                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
'                Select Case RSTtax!DISC_FLAG
'                    Case "A"
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                    Case Else
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                End Select
'                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n)) * TAX_PER / 100
            GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
            n = n + 2
        Loop
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND CATEGORY = 'SERVICE CHARGE'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            'GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
            GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
            TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)) * TAX_PER / 100
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)), 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 1)), 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Format(Round(CESSPER, 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Format(Round(CESSAMT, 3), "0.00")
        
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)) * TAX_PER / 100
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)), 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 3)), 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 6) = Format(Round(CESSPER, 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 5) = Format(Round(CESSAMT, 3), "0.00")
        
        n = 6
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        M = M + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    For i = 6 To GRDTranx.Cols - 4
        GRDTranx.TextMatrix(0, i) = "Purchase " & GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    'db.Execute "create table TEMP_REPORT (Col1 number, Col2 Varchar(15))"
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Sales_Register_Sum()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX, rstTRANX_DATE As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim NETAMOUNT, TAXSALE1, TAXAMT1, TAXAMT2, TAXSALE2, TAXAMT3, TAXSALE3, TAXAMT4, TAXSALE4, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    
    Set rstTRANX = New ADODB.Recordset
    If OPTGST.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf OptGR.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf Optservice.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf OptRT.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf OptWs.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf Optst.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    ElseIf Opt8V.Value = True Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='VI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    End If
    GRDTranx.rows = 1
    GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Customer."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 2000
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    GrdTotal.ColAlignment(5) = 4
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If GRDTranx.rows = 6 Then Exit Function

    n = 6
    M = 1
    
    Dim COUNT As Long
    Set rstTRANX_DATE = New ADODB.Recordset
    If OPTGST.Value = True Then
        rstTRANX_DATE.Open "SELECT DISTINCT VCH_DATE From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY VCH_DATE", db, adOpenStatic, adLockReadOnly
    ElseIf OptRT.Value = True Then
        rstTRANX_DATE.Open "SELECT DISTINCT VCH_DATE From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY VCH_DATE", db, adOpenStatic, adLockReadOnly
    ElseIf OptWs.Value = True Then
        rstTRANX_DATE.Open "SELECT DISTINCT VCH_DATE From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY VCH_DATE", db, adOpenStatic, adLockReadOnly
    End If
    Do Until rstTRANX_DATE.EOF
        NETAMOUNT = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(rstTRANX_DATE!VCH_DATE) & "' AND TRX_TYPE= '" & rstTRANX_DATE!TRX_TYPE & "'  ", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            NETAMOUNT = NETAMOUNT + IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            Do Until n = GRDTranx.Cols
                COUNT = 1
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100)
                    Else
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                n = n + 2
                COUNT = COUNT + 1
            Loop
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
                
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 4) = rstTRANX!NET_AMOUNT
        GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        rstTRANX_DATE.MoveNext
    Loop
    rstTRANX_DATE.Close
    Set rstTRANX_DATE = New ADODB.Recordset
    For i = 6 To GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Sales_Register_CST()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    
    Set rstTRANX = New ADODB.Recordset
    GRDTranx.rows = 1
    GRDTranx.Cols = 9
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Customer."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.TextMatrix(0, 6) = "CST %"
    GRDTranx.TextMatrix(0, 7) = "Tax Sale Amount"
    GRDTranx.TextMatrix(0, 8) = "Tax Amount"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1800
    GRDTranx.ColWidth(6) = 1500
    GRDTranx.ColWidth(7) = 2000
    GRDTranx.ColWidth(8) = 1800
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1800
    GrdTotal.ColWidth(6) = 1500
    GrdTotal.ColWidth(7) = 2000
    GrdTotal.ColWidth(8) = 1800
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    GrdTotal.ColAlignment(5) = 4
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If OPTGST.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        ElseIf OptGR.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Optservice.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        Else
            BIL_PRE = ""
            BILL_SUF = ""
        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    n = 6
    M = 1
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE CST > 0 AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
        GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
        GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
        GRDTranx.TextMatrix(M, 4) = rstTRANX!NET_AMOUNT
        GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!CST), 0, rstTRANX!CST)
        GRDTranx.TextMatrix(M, 7) = IIf(IsNull(rstTRANX!NET_AMOUNT), "", rstTRANX!NET_AMOUNT)
        GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 7)) * 100 / (Val(GRDTranx.TextMatrix(M, 6)) + 100), 3), "0.00")
        GRDTranx.TextMatrix(M, 8) = Format(Round((Val(GRDTranx.TextMatrix(M, 7)) * GRDTranx.TextMatrix(M, 6)) / 100, 3), "0.00")
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        M = M + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing

    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Sub OptSalesreturn_Click()
    FrameSales.Visible = False
    FrmPurchase.Visible = False
    Frmeperiod.Caption = "Sales Return Register"
End Sub

Private Function PURCH_RET_Register()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PR') ", db, adOpenStatic, adLockReadOnly

    GRDTranx.rows = 1
    GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Customer."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1300
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    GrdTotal.ColAlignment(5) = 4
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If GRDTranx.rows = 6 Then Exit Function
    
    n = 6
    M = 1
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From PURCAHSERETURN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PR')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
        GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
        GRDTranx.TextMatrix(M, 3) = Format(rstTRANX!VCH_NO, bill_for)
        GRDTranx.TextMatrix(M, 4) = rstTRANX!NET_AMOUNT
        GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        Do Until n = GRDTranx.Cols
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & "", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                Select Case rstTRANX!SLSM_CODE
                    Case "P"
                        GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                End Select
                If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                    GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                Else
                    GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                End If
                
                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
            n = n + 2
        Loop
        n = 6
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        M = M + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    For i = 6 To GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, i) = GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function SALES_RET_REGISTER()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim TAX_PER As Single
    Dim RSTACTMAST As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim TOTAL_AMT As Double
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    On Error GoTo ERRHAND
    
    If MDIMAIN.lblgst.Caption = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly
        'rstTRANX.Open "Select DISTINCT SALES_TAX From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2 + 1
        GrdTotal.Cols = GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer"
        GRDTranx.TextMatrix(0, 2) = "GSTin No"
        GRDTranx.TextMatrix(0, 3) = "Credit Note No."
        GRDTranx.TextMatrix(0, 4) = "Amount"
        GRDTranx.TextMatrix(0, 5) = "Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3200
        GRDTranx.ColWidth(2) = 1700
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1400
        GRDTranx.ColWidth(5) = 1100
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3200
        GrdTotal.ColWidth(2) = 1700
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1400
        GrdTotal.ColWidth(5) = 1100
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Invoice Details"
        GRDTranx.ColAlignment(GRDTranx.Cols - 1) = 1
        GRDTranx.ColWidth(GRDTranx.Cols - 1) = 3000
        If GRDTranx.rows = 7 Then Exit Function
        
        n = 6
        M = 1
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            TOTAL_AMT = 0
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            
            Set RSTACTMAST = New ADODB.Recordset
            RSTACTMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
            End If
            RSTACTMAST.Close
            Set RSTACTMAST = Nothing
            
            GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!VCH_NO), "", "SR-" & Format(rstTRANX!VCH_NO, bill_for))
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!VCH_AMOUNT), "", rstTRANX!VCH_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Do Until n = GRDTranx.Cols - 1
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE RTRXFILE.TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND RTRXFILE.VCH_NO = " & rstTRANX!VCH_NO & " AND RTRXFILE.TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND RTRXFILE.SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y')", db, adOpenStatic, adLockReadOnly, adCmdText
                'RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & "", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
    '                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    '                Select Case RSTtax!DISC_FLAG
    '                    Case "A"
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                    Case Else
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                End Select
    '                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n)) * TAX_PER / 100
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT, 3), "0.00")
            GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = IIf(IsNull(rstTRANX!INV_DETAILS), "", rstTRANX!INV_DETAILS)
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        For i = 6 To GRDTranx.Cols - 2
            GRDTranx.TextMatrix(0, i) = GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 7
        GrdTotal.Cols = GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer"
        GRDTranx.TextMatrix(0, 2) = "GSTin No"
        GRDTranx.TextMatrix(0, 3) = "Credit Note No."
        GRDTranx.TextMatrix(0, 4) = "Amount"
        GRDTranx.TextMatrix(0, 5) = "Date"
        GRDTranx.TextMatrix(0, 6) = "Bill Details"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1900
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 2000
        GRDTranx.ColWidth(6) = 3500
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GRDTranx.ColAlignment(5) = 4
        GRDTranx.ColAlignment(6) = 1
        
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        
        
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            
            Set RSTACTMAST = New ADODB.Recordset
            RSTACTMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
            End If
            RSTACTMAST.Close
            Set RSTACTMAST = Nothing
            
            GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!VCH_NO), "", "SR-" & Format(rstTRANX!VCH_NO, bill_for))
            GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!VCH_AMOUNT), "", rstTRANX!VCH_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!INV_DETAILS), "", rstTRANX!INV_DETAILS)
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
    End If
    
    'db.Execute "create table TEMP_REPORT (Col1 number, Col2 Varchar(15))"
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GrdTotal.rows = 0
    GrdTotal.rows = GrdTotal.rows + 1
    GrdTotal.Cols = GRDTranx.Cols
    GrdTotal.TextMatrix(0, 3) = "TOTAL"
    For n = 4 To GRDTranx.Cols - 2
        If n <> 5 Then
            For i = 1 To GRDTranx.rows - 1
                GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
            Next i
        End If
    Next n
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Sales_Register_Daily()
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If OPTGST.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        ElseIf OptGR.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Optservice.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        Else
            BIL_PRE = ""
            BILL_SUF = ""
        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim CESSPER As Double
    Dim CESSAMT As Double
    
    Dim FIRST_BILL As Double
    Dim LAST_BILL As Double
    Dim FROMDATE As Date
    Dim TODATE As Date
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optst.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='VI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        End If
        GRDTranx.rows = 1
        GRDTranx.Cols = 9 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = ""
        GRDTranx.TextMatrix(0, 2) = ""
        GRDTranx.TextMatrix(0, 3) = "BILL NOS"
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 0 '3500
        GRDTranx.ColWidth(2) = 0 '1800
        GRDTranx.ColWidth(3) = 3500
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 0 '3500
        GrdTotal.ColWidth(2) = 0 '1800
        GrdTotal.ColWidth(3) = 3500
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        n = 6
        M = 1
                
        Dim TOTAL_AMT As Double
        Dim KFC As Double
        Dim DISC_AMT As Double
        FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
        TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
        Do Until FROMDATE > TODATE
            Set rstTRANX = New ADODB.Recordset
            If OPTGST.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptGR.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Optservice.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Optst.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptRT.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Opt8V.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptWs.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' ) ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            End If
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                TOTAL_AMT = 0
                CESSPER = 0
                CESSAMT = 0
                KFC = 0
                rstTRANX.MoveLast
                LAST_BILL = rstTRANX!VCH_NO
                rstTRANX.MoveFirst
                FIRST_BILL = rstTRANX!VCH_NO
                GRDTranx.rows = GRDTranx.rows + 1
                GRDTranx.FixedRows = 1
                GRDTranx.TextMatrix(M, 0) = M
                GRDTranx.TextMatrix(M, 1) = "" 'IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                GRDTranx.TextMatrix(M, 2) = "" 'IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
                GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(FIRST_BILL, "0000") & BILL_SUF & " TO " & BIL_PRE & Format(LAST_BILL, "0000") & BILL_SUF
                GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
                Do Until n = GRDTranx.Cols - 3
                    Set RSTtax = New ADODB.Recordset
                    RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                    Do Until RSTtax.EOF
                        'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                        'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)\
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                            Select Case RSTTRXFILE!SLSM_CODE
'                                Case "P"
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
'                                Case Else
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
'                            End Select
'                            GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                            
                            Select Case RSTTRXFILE!SLSM_CODE
                                Case "P"
                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                                Case Else
                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                            End Select
                            If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                                KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                                CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                            Else
                                KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                                CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                            End If
                            
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                        'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                            GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        Else
                            GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        End If
                        'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        RSTtax.MoveNext
                    Loop
                    RSTtax.Close
                    Set RSTtax = Nothing
                    
                    GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                    GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                    TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                    n = n + 2
                Loop
                GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
                GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
                GRDTranx.TextMatrix(M, 4) = TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                
                DISC_AMT = 0
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    If RSTtax!SLSM_CODE = "A" Then
                        DISC_AMT = DISC_AMT + IIf(IsNull(RSTtax!DISCOUNT), 0, RSTtax!DISCOUNT)
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
                
                n = 6
                vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
                M = M + 1
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            FROMDATE = DateAdd("d", FROMDATE, 1)
        Loop
        For i = 6 To GRDTranx.Cols - 1
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "KFC"
        
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 4
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "BILL NOS"
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 2) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 2) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 3) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1500
        GRDTranx.ColWidth(3) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1500
        GrdTotal.ColWidth(3) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 4
        GRDTranx.ColAlignment(2) = 4
        GRDTranx.ColAlignment(3) = 4
        GrdTotal.ColAlignment(0) = 4
        GrdTotal.ColAlignment(1) = 4
        GrdTotal.ColAlignment(2) = 4
        GrdTotal.ColAlignment(3) = 4

        n = 6
        M = 1
        
        FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
        TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
        Do Until FROMDATE > TODATE
            Set rstTRANX = New ADODB.Recordset
            If OPTGST.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptGR.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Optservice.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Optst.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptRT.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf Opt8V.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            ElseIf OptWs.Value = True Then
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' ) ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI' or TRX_TYPE='GI' or TRX_TYPE='SV') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            End If
            
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                rstTRANX.MoveLast
                LAST_BILL = rstTRANX!VCH_NO
                rstTRANX.MoveFirst
                FIRST_BILL = rstTRANX!VCH_NO
                GRDTranx.rows = GRDTranx.rows + 1
                GRDTranx.FixedRows = 1
                GRDTranx.TextMatrix(M, 0) = M
                GRDTranx.TextMatrix(M, 1) = BIL_PRE & Format(FIRST_BILL, "0000") & BILL_SUF & " TO " & BIL_PRE & Format(LAST_BILL, "0000") & BILL_SUF
                GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    GRDTranx.TextMatrix(M, 2) = Val(GRDTranx.TextMatrix(M, 2)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                DISC_AMT = 0
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXMAST WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    DISC_AMT = DISC_AMT + IIf(IsNull(RSTtax!DISCOUNT), 0, RSTtax!DISCOUNT)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, 2) = Format(Round(Val(GRDTranx.TextMatrix(M, 2)) - DISC_AMT, 2), "0.00")
                
                vbalProgressBar1.Value = vbalProgressBar1.Value + 1
                M = M + 1
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            FROMDATE = DateAdd("d", FROMDATE, 1)
        Loop
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 2 To GRDTranx.Cols - 1
            If n <> 3 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Assets_Register(BillType As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    On Error GoTo ERRHAND
    
    Set rstTRANX = New ADODB.Recordset
    If BillType = 1 Then
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From ASTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='AP') ", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From ASTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='EP') ", db, adOpenStatic, adLockReadOnly
    End If
    GRDTranx.rows = 1
    GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Supplier."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1500
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    If GRDTranx.rows = 6 Then Exit Function
    
    n = 6
    M = 1
    Dim TAX_PER As Single
    Dim RSTACTMAST As ADODB.Recordset
    Set rstTRANX = New ADODB.Recordset
    If BillType = 1 Then
        rstTRANX.Open "SELECT * From ASTRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='AP')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From ASTRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='EP')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!ACT_NAME
        
        Set RSTACTMAST = New ADODB.Recordset
        RSTACTMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
        End If
        RSTACTMAST.Close
        Set RSTACTMAST = Nothing
        
        GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!PINV), "", rstTRANX!PINV)
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), "", rstTRANX!NET_AMOUNT)
        GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        Do Until n = GRDTranx.Cols
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From ASTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & "", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
'                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
'                Select Case RSTtax!DISC_FLAG
'                    Case "A"
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                    Case Else
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                End Select
'                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n)) * TAX_PER / 100
            GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
            n = n + 2
        Loop
        n = 6
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        M = M + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    For i = 6 To GRDTranx.Cols - 1
        GRDTranx.TextMatrix(0, i) = "Purchase " & GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    'db.Execute "create table TEMP_REPORT (Col1 number, Col2 Varchar(15))"
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function


Private Function Sales_Register_GST()
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If OPTGST.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
        ElseIf OptGR.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        ElseIf Optservice.Value = True Then
            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
        Else
            BIL_PRE = ""
            BILL_SUF = ""
        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Optst.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='TF' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='VI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        End If
        GRDTranx.rows = 1
        GRDTranx.Cols = 8 + rstTRANX.RecordCount * 4
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX / 2 & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 4
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 8 Then Exit Function
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            TOTAL_AMT = 0
            Do Until n = GRDTranx.Cols - 2
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Addl Compensation Cess"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        If OPTGST.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptGR.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='HI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Optservice.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptRT.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='RI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf Opt8V.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptWs.Value = True Then
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SI' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='TF' )  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, 4) = Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing

            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function EX_RET_REGISTER()
    
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim RSTACTMAST As ADODB.Recordset
    
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim CUSTCODE As String
    Dim TAX_PER As Single
    
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    On Error GoTo ERRHAND
    
    If MDIMAIN.lblgst.Caption = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer"
        GRDTranx.TextMatrix(0, 2) = "GSTin No"
        GRDTranx.TextMatrix(0, 3) = "No."
        GRDTranx.TextMatrix(0, 4) = "Amount"
        GRDTranx.TextMatrix(0, 5) = "Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 2000
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GRDTranx.ColAlignment(5) = 4
        
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 6 Then Exit Function
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT TRX_YEAR, TRX_TYPE, VCH_NO From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV')  ORDER BY TRX_TYPE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            'GRDTranx.TextMatrix(M, 1) = "CASH" 'rstTRANX!ACT_NAME
            'GRDTranx.TextMatrix(M, 2) = ""
            GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!VCH_NO), "", Format(rstTRANX!VCH_NO, bill_for))
            CUSTCODE = ""
            Do Until n = GRDTranx.Cols
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & "", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
    '                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    '                Select Case RSTtax!DISC_FLAG
    '                    Case "A"
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                    Case Else
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                End Select
    '                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                    
                    GRDTranx.TextMatrix(M, 5) = IIf(IsDate(RSTtax!VCH_DATE), RSTtax!VCH_DATE, "")
                    GRDTranx.TextMatrix(M, 4) = Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), "", RSTtax!TRX_TOTAL)
                    If Not (IsNull(RSTtax!M_USER_ID) Or RSTtax!M_USER_ID = "") Then
                        CUSTCODE = RSTtax!M_USER_ID
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n)) * TAX_PER / 100
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                
                Set RSTACTMAST = New ADODB.Recordset
                RSTACTMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & CUSTCODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(RSTACTMAST!ACT_NAME), "", RSTACTMAST!ACT_NAME)
                    GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
                Else
                    GRDTranx.TextMatrix(M, 1) = "CASH"
                    GRDTranx.TextMatrix(M, 2) = ""
                End If
                RSTACTMAST.Close
                Set RSTACTMAST = Nothing
                
                n = n + 2
            Loop
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 1
            GRDTranx.TextMatrix(0, i) = GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
    Else
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 6 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer"
        GRDTranx.TextMatrix(0, 2) = "GSTin No"
        GRDTranx.TextMatrix(0, 3) = "No."
        GRDTranx.TextMatrix(0, 4) = "Amount"
        GRDTranx.TextMatrix(0, 5) = "Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 2000
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GRDTranx.ColAlignment(5) = 4
        
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 6 Then Exit Function
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT TRX_YEAR, TRX_TYPE, VCH_NO From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            'GRDTranx.TextMatrix(M, 1) = "CASH" 'rstTRANX!ACT_NAME
            'GRDTranx.TextMatrix(M, 2) = ""
            GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!VCH_NO), "", Format(rstTRANX!VCH_NO, bill_for))
            CUSTCODE = ""
            Do Until n = GRDTranx.Cols
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & "", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
    '                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
    '                Select Case RSTtax!DISC_FLAG
    '                    Case "A"
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                    Case Else
    '                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
    '                End Select
    '                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                    
                    GRDTranx.TextMatrix(M, 5) = IIf(IsDate(RSTtax!VCH_DATE), RSTtax!VCH_DATE, "")
                    GRDTranx.TextMatrix(M, 4) = Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), "", RSTtax!TRX_TOTAL)
                    If Not (IsNull(RSTtax!M_USER_ID) Or RSTtax!M_USER_ID = "") Then
                        CUSTCODE = RSTtax!M_USER_ID
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N)) * TAX_PER / 100
                'GRDTranx.TextMatrix(M, N) = Format(Round(Val(GRDTranx.TextMatrix(M, N)), 3), "0.00")
                'GRDTranx.TextMatrix(M, N + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, N + 1)), 3), "0.00")
                
                Set RSTACTMAST = New ADODB.Recordset
                RSTACTMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & CUSTCODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(RSTACTMAST!ACT_NAME), "", RSTACTMAST!ACT_NAME)
                    GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
                Else
                    GRDTranx.TextMatrix(M, 1) = "CASH"
                    GRDTranx.TextMatrix(M, 2) = ""
                End If
                RSTACTMAST.Close
                Set RSTACTMAST = Nothing
                
                n = n + 2
            Loop
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 1
            GRDTranx.TextMatrix(0, i) = ""
        Next
        GRDTranx.Cols = 6
    End If
    'db.Execute "create table TEMP_REPORT (Col1 number, Col2 Varchar(15))"
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Sales_RegisterBR()
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
'        If OptGST.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
'        ElseIf OptGR.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
'        ElseIf OptService.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
'        Else
'            BIL_PRE = ""
'            BILL_SUF = ""
'        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILEVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 10 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                Else
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Or rstTRANX!BILL_NAME = "CASH" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            'Trim(txtBillNo.text)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            Dim KFC As Double
            TOTAL_AMT = 0
            KFC = 0
            Do Until n = GRDTranx.Cols - 4
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILEVAN WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 3)
            GRDTranx.TextMatrix(M, n + 3) = Format(Round(DISC_AMT, 3), "0.00")
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 2
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "KFC"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "DISC"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 4) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMASTVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(rstTRANX!VCH_NO, bill_for) & BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILEVAN WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL), 3)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing

            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Sales_Register_DailyBR()
    
    Dim BIL_PRE, BILL_SUF, GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
'        If OptGST.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
'        ElseIf OptGR.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
'        ElseIf OptService.value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
'        Else
'            BIL_PRE = ""
'            BILL_SUF = ""
'        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim CESSPER As Double
    Dim CESSAMT As Double
    
    Dim FIRST_BILL As Double
    Dim LAST_BILL As Double
    Dim FROMDATE As Date
    Dim TODATE As Date
    
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILEVAN WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        
        GRDTranx.rows = 1
        GRDTranx.Cols = 9 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = ""
        GRDTranx.TextMatrix(0, 2) = ""
        GRDTranx.TextMatrix(0, 3) = "BILL NOS"
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 0 '3500
        GRDTranx.ColWidth(2) = 0 '1800
        GRDTranx.ColWidth(3) = 3500
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 0 '3500
        GrdTotal.ColWidth(2) = 0 '1800
        GrdTotal.ColWidth(3) = 3500
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        n = 6
        M = 1
                
        Dim TOTAL_AMT As Double
        Dim KFC As Double
        Dim DISC_AMT As Double
        FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
        TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
        Do Until FROMDATE > TODATE
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From TRXMASTVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                TOTAL_AMT = 0
                CESSPER = 0
                CESSAMT = 0
                KFC = 0
                rstTRANX.MoveLast
                LAST_BILL = rstTRANX!VCH_NO
                rstTRANX.MoveFirst
                FIRST_BILL = rstTRANX!VCH_NO
                GRDTranx.rows = GRDTranx.rows + 1
                GRDTranx.FixedRows = 1
                GRDTranx.TextMatrix(M, 0) = M
                GRDTranx.TextMatrix(M, 1) = "" 'IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                GRDTranx.TextMatrix(M, 2) = "" 'IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
                GRDTranx.TextMatrix(M, 3) = BIL_PRE & Format(FIRST_BILL, "0000") & BILL_SUF & " TO " & BIL_PRE & Format(LAST_BILL, "0000") & BILL_SUF
                GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
                Do Until n = GRDTranx.Cols - 3
                    Set RSTtax = New ADODB.Recordset
                    RSTtax.Open "Select * From TRXFILEVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                    Do Until RSTtax.EOF
                        'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                        'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)\
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "SELECT * From TRXMASTVAN WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                            Select Case RSTTRXFILE!SLSM_CODE
'                                Case "P"
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
'                                Case Else
'                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
'                            End Select
'                            GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                            
                            Select Case RSTTRXFILE!SLSM_CODE
                                Case "P"
                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                                Case Else
                                    GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                            End Select
                            If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                                KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                                CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                            Else
                                KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                                CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                            End If
                            'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                            
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                        'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                            GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        Else
                            GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        End If
                        'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        RSTtax.MoveNext
                    Loop
                    RSTtax.Close
                    Set RSTtax = Nothing
                    
                    GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                    GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                    TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                    n = n + 2
                Loop
                GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
                GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
                GRDTranx.TextMatrix(M, 4) = TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                
                DISC_AMT = 0
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXMASTVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    If RSTtax!SLSM_CODE = "A" Then
                        DISC_AMT = DISC_AMT + IIf(IsNull(RSTtax!DISCOUNT), 0, RSTtax!DISCOUNT)
                    End If
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
                
                n = 6
                vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
                M = M + 1
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            FROMDATE = DateAdd("d", FROMDATE, 1)
        Loop
        For i = 6 To GRDTranx.Cols - 1
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "KFC"
        
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 4
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "BILL NOS"
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 2) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 2) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 3) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1500
        GRDTranx.ColWidth(3) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1500
        GrdTotal.ColWidth(3) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 4
        GRDTranx.ColAlignment(2) = 4
        GRDTranx.ColAlignment(3) = 4
        GrdTotal.ColAlignment(0) = 4
        GrdTotal.ColAlignment(1) = 4
        GrdTotal.ColAlignment(2) = 4
        GrdTotal.ColAlignment(3) = 4

        n = 6
        M = 1
        
        FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
        TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
        Do Until FROMDATE > TODATE
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From TRXMASTVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
            
            If Not (rstTRANX.EOF And rstTRANX.BOF) Then
                rstTRANX.MoveLast
                LAST_BILL = rstTRANX!VCH_NO
                rstTRANX.MoveFirst
                FIRST_BILL = rstTRANX!VCH_NO
                GRDTranx.rows = GRDTranx.rows + 1
                GRDTranx.FixedRows = 1
                GRDTranx.TextMatrix(M, 0) = M
                GRDTranx.TextMatrix(M, 1) = BIL_PRE & Format(FIRST_BILL, "0000") & BILL_SUF & " TO " & BIL_PRE & Format(LAST_BILL, "0000") & BILL_SUF 'BIL_PRE & FIRST_BILL & BILL_SUF & " TO " & BIL_PRE & LAST_BILL & BILL_SUF
                GRDTranx.TextMatrix(M, 3) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILEVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    GRDTranx.TextMatrix(M, 2) = Val(GRDTranx.TextMatrix(M, 2)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                DISC_AMT = 0
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXMASTVAN WHERE VCH_DATE ='" & Format(FROMDATE, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    DISC_AMT = DISC_AMT + IIf(IsNull(RSTtax!DISCOUNT), 0, RSTtax!DISCOUNT)
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                GRDTranx.TextMatrix(M, 2) = Format(Round(Val(GRDTranx.TextMatrix(M, 2)) - DISC_AMT, 2), "0.00")
                
                vbalProgressBar1.Value = vbalProgressBar1.Value + 1
                M = M + 1
            End If
            rstTRANX.Close
            Set rstTRANX = Nothing
            FROMDATE = DateAdd("d", FROMDATE, 1)
        Loop
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 2 To GRDTranx.Cols - 1
            If n <> 3 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Purchase_Register_GSTR1()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    'On Error Resume Next
    'db.Execute "DROP TABLE TEMP_REPORT "
    'On Error GoTo ErrHand
    
    Set rstTRANX = New ADODB.Recordset
    If OptNormal.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    ElseIf OptComm.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP') AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly
        End If
    End If
    GRDTranx.rows = 1
    GRDTranx.Cols = (6 + rstTRANX.RecordCount * 2) + 4 + 1 + 1
    GrdTotal.Cols = GRDTranx.Cols
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "Supplier."
    GRDTranx.TextMatrix(0, 2) = "GSTin No"
    GRDTranx.TextMatrix(0, 3) = "Bill No."
    GRDTranx.TextMatrix(0, 4) = "Bill Amt"
    GRDTranx.TextMatrix(0, 5) = "Bill Date"
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3500
    GRDTranx.ColWidth(2) = 1800
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 1300
    
    GrdTotal.ColWidth(0) = 800
    GrdTotal.ColWidth(1) = 3500
    GrdTotal.ColWidth(2) = 1800
    GrdTotal.ColWidth(3) = 1100
    GrdTotal.ColWidth(4) = 1500
    GrdTotal.ColWidth(5) = 1500
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GrdTotal.ColAlignment(3) = 4
    GrdTotal.ColAlignment(4) = 4
    
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "Ser. Taxable Amt"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Ser Tax Amt"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess"
'    GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Cess"
    
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Ser. Taxable Amt"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Ser Tax Amt"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 6) = "Cess"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 5) = "Addl Cess"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "Discount"
    GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "TCS"
                    
'    GrdTotal.ColWidth(GRDTranx.Cols - 2) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 1) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 4) = 1500
'    GrdTotal.ColWidth(GRDTranx.Cols - 3) = 1500
    
    GrdTotal.ColWidth(GRDTranx.Cols - 4) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 3) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 5) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 6) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 2) = 1500
    GrdTotal.ColWidth(GRDTranx.Cols - 1) = 1500
    
    GrdTotal.ColAlignment(GRDTranx.Cols - 1) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 2) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 6) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 4) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 3) = 4
    GrdTotal.ColAlignment(GRDTranx.Cols - 5) = 4
                    
    i = 6
    Do Until rstTRANX.EOF
        GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
        GRDTranx.ColWidth(i) = 1600
        GRDTranx.ColAlignment(i) = 4
        GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
        GRDTranx.ColWidth(i + 1) = 1600
        GRDTranx.ColAlignment(i + 1) = 4
        
        GrdTotal.ColWidth(i) = 1600
        GrdTotal.ColAlignment(i) = 4
        GrdTotal.ColWidth(i + 1) = 1600
        GrdTotal.ColAlignment(i + 1) = 4
        
        i = i + 2
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If GRDTranx.rows = 6 Then Exit Function
    
    n = 6
    M = 1
    Dim TAX_PER As Single
    Dim CESSPER As Double
    Dim CESSAMT As Double
    
    Dim RSTACTMAST As ADODB.Recordset
    Set rstTRANX = New ADODB.Recordset
    If OptNormal.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    ElseIf OptComm.Value = True Then
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PW')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If Optinvdate.Value = True Then
            rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRANSMAST WHERE RCVD_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCVD_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='LP')  ORDER BY TRX_TYPE, RCVD_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    End If
    Do Until rstTRANX.EOF
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!ACT_NAME
        
        Set RSTACTMAST = New ADODB.Recordset
        RSTACTMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
        End If
        RSTACTMAST.Close
        Set RSTACTMAST = Nothing
        
        GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!PINV), "", rstTRANX!PINV)
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), "", rstTRANX!NET_AMOUNT)
        If Optinvdate.Value = True Then
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
        Else
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!RCVD_DATE), rstTRANX!RCVD_DATE, "")
        End If
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = IIf(IsNull(rstTRANX!ADD_AMOUNT), 0, rstTRANX!ADD_AMOUNT)
        CESSPER = 0
        CESSAMT = 0
        Do Until n = GRDTranx.Cols - 4
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & Val(GRDTranx.TextMatrix(0, n)) & " AND (ISNULL(CATEGORY) OR UCASE(CATEGORY) <> 'SERVICE CHARGE' )", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                
                Select Case RSTtax!DISC_FLAG
                    Case "P"
                        GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!TR_DISC / 100) '- ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                    Case Else
                        GRDTranx.Tag = RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY)) - ((RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY))) * RSTtax!TR_DISC / 100)
                End Select
                GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
                GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))


                TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
                If RSTtax!DISC_FLAG = "P" Then
                    CESSPER = CESSPER + (Val(GRDTranx.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                Else
                    CESSPER = CESSPER + (Val(GRDTranx.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                End If
                CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
'                'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
'                Select Case RSTtax!DISC_FLAG
'                    Case "A"
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - RSTtax!P_DISC) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                    Case Else
'                        GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
'                End Select
'                'GRDTranx.TextMatrix(M, N + 1) = Val(GRDTranx.TextMatrix(M, N + 1)) + Val(GRDTranx.TextMatrix(M, N)) * RSTtax!SALES_TAX / 100
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n)) * TAX_PER / 100
            GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
            n = n + 2
        Loop
        Set RSTtax = New ADODB.Recordset
        RSTtax.Open "Select * From RTRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND CATEGORY = 'SERVICE CHARGE'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTtax.EOF
            'GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
            GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
            TAX_PER = IIf(IsNull(RSTtax!SALES_TAX), 0, RSTtax!SALES_TAX)
            RSTtax.MoveNext
        Loop
        RSTtax.Close
        Set RSTtax = Nothing
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)) * TAX_PER / 100
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 2) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 2)), 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 1) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 1)), 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Format(Round(CESSPER, 3), "0.00")
'        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Format(Round(CESSAMT, 3), "0.00")
        
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)) * TAX_PER / 100
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 4) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 4)), 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 3) = Format(Round(Val(GRDTranx.TextMatrix(M, GRDTranx.Cols - 3)), 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 6) = Format(Round(CESSPER, 3), "0.00")
        GRDTranx.TextMatrix(M, GRDTranx.Cols - 5) = Format(Round(CESSAMT, 3), "0.00")
        
        n = 6
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        M = M + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    For i = 6 To GRDTranx.Cols - 4
        GRDTranx.TextMatrix(0, i) = "Purchase " & GRDTranx.TextMatrix(0, i) & "%"
        i = i + 1
    Next
    
    'db.Execute "create table TEMP_REPORT (Col1 number, Col2 Varchar(15))"
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function Damage_Register()
    
    Dim GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
'        If OptGST.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
'        ElseIf OptGR.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
'        ElseIf OptService.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
'        Else
'            BIL_PRE = ""
'            BILL_SUF = ""
'        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='DM' OR TRX_TYPE='DG') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 10 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DAMAGE_MAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='DM' OR TRX_TYPE='DG') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                Else
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Or rstTRANX!BILL_NAME = "CASH" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            'Trim(txtBillNo.text)
            GRDTranx.TextMatrix(M, 3) = Format(rstTRANX!VCH_NO, bill_for) '& BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            Dim KFC As Double
            TOTAL_AMT = 0
            KFC = 0
            Do Until n = GRDTranx.Cols - 4
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY

                    End If
                    
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 3)
            GRDTranx.TextMatrix(M, n + 3) = Format(Round(DISC_AMT, 3), "0.00")
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 2
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "KFC"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "DISC"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 4) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DAMAGE_MAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='DM' OR TRX_TYPE='DG') ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = Format(rstTRANX!VCH_NO, bill_for) '& BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL), 3)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing

            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

Private Function QTN_Register()
    
    Dim GST_FLAG As String
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
'        If OptGST.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8), "", RSTCOMPANY!SUFIX_8)
'        ElseIf OptGR.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
'        ElseIf OptService.Value = True Then
'            BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8B)
'            BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8B), "", RSTCOMPANY!SUFIX_8B)
'        Else
'            BIL_PRE = ""
'            BILL_SUF = ""
'        End If
        GST_FLAG = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim i As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    Dim DISC_AMT As Double
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    'If GST_FLAG = "R" And Optst.value = False Then
    If GST_FLAG = "R" Then
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT DISTINCT SALES_TAX From QTNSUB WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='QT' ", db, adOpenStatic, adLockReadOnly
        GRDTranx.rows = 1
        GRDTranx.Cols = 10 + rstTRANX.RecordCount * 2
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        i = 6
        Do Until rstTRANX.EOF
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX
            GRDTranx.ColWidth(i) = 1600
            GRDTranx.ColAlignment(i) = 4
            GRDTranx.TextMatrix(0, i + 1) = "Tax Amt " & rstTRANX!SALES_TAX & "%"
            GRDTranx.ColWidth(i + 1) = 1600
            GRDTranx.ColAlignment(i + 1) = 4
            
            GrdTotal.ColWidth(i) = 1600
            GrdTotal.ColAlignment(i) = 4
            GrdTotal.ColWidth(i + 1) = 1600
            GrdTotal.ColAlignment(i + 1) = 4
            
            i = i + 2
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        If GRDTranx.rows = 9 Then Exit Function
        
        n = 6
        M = 1
        Dim CESSPER As Double
        Dim CESSAMT As Double
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From QTNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='QT' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                Else
                    If IsNull(rstTRANX!BILL_NAME) Or rstTRANX!BILL_NAME = "" Or rstTRANX!BILL_NAME = "CASH" Then
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    Else
                        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
                    End If
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS) Or rstTRANX!BILL_ADDRESS = "", "", ", " & rstTRANX!BILL_ADDRESS)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            'Trim(txtBillNo.text)
            GRDTranx.TextMatrix(M, 3) = Format(rstTRANX!VCH_NO, bill_for) '& BILL_SUF
            'GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            CESSPER = 0
            CESSAMT = 0
            Dim TOTAL_AMT As Double
            Dim KFC As Double
            TOTAL_AMT = 0
            KFC = 0
            Do Until n = GRDTranx.Cols - 4
                Set RSTtax = New ADODB.Recordset
                RSTtax.Open "Select * From QTNSUB WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & GRDTranx.TextMatrix(0, n) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTtax.EOF
                    'GRDTranx.TextMatrix(M, N) = Val(GRDTranx.TextMatrix(M, N)) + (RSTtax!TRX_TOTAL * 100) / ((RSTtax!SALES_TAX) + 100)
                    'TXTRETAILNOTAX.Text = Round(Val(TXTRETAIL.Text) * 100 / (Val(TXTTAX.Text) + 100), 4)
                    Select Case rstTRANX!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    If IsNull(RSTtax!QTY) Or RSTtax!QTY = 0 Then
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100))
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100)
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt)
                    Else
                        KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!kfc_tax), 0, RSTtax!kfc_tax / 100)) * RSTtax!QTY
                        GRDTranx.TextMatrix(M, n) = Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                        GRDTranx.TextMatrix(M, n + 1) = Val(GRDTranx.TextMatrix(M, n + 1)) + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                        CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                        CESSAMT = CESSAMT + IIf(IsNull(RSTtax!cess_amt), 0, RSTtax!cess_amt) * RSTtax!QTY
                    End If
                    
                    RSTtax.MoveNext
                Loop
                RSTtax.Close
                Set RSTtax = Nothing
                
                GRDTranx.TextMatrix(M, n) = Format(Round(Val(GRDTranx.TextMatrix(M, n)), 3), "0.00")
                GRDTranx.TextMatrix(M, n + 1) = Format(Round(Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                TOTAL_AMT = TOTAL_AMT + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1))
                n = n + 2
            Loop
            GRDTranx.TextMatrix(M, n) = Format(Round(CESSPER, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 1) = Format(Round(CESSAMT, 3), "0.00")
            GRDTranx.TextMatrix(M, n + 2) = Format(Round(KFC, 3), "0.00")
            GRDTranx.TextMatrix(M, 4) = Format(Round(TOTAL_AMT + KFC + Val(GRDTranx.TextMatrix(M, n)) + Val(GRDTranx.TextMatrix(M, n + 1)), 3), "0.00")
                        
            DISC_AMT = 0
            If rstTRANX!SLSM_CODE = "A" Then
                DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            End If
            GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 3)
            GRDTranx.TextMatrix(M, n + 3) = Format(Round(DISC_AMT, 3), "0.00")
            n = 6
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        For i = 6 To GRDTranx.Cols - 3
            GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
            i = i + 1
        Next
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 2
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 4) = "Cess Amount"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 3) = "Addl Compensation Cess"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 2) = "KFC"
        GRDTranx.TextMatrix(0, GRDTranx.Cols - 1) = "DISC"
    Else
        GRDTranx.rows = 1
        GRDTranx.Cols = 6
        GrdTotal.Cols = GRDTranx.Cols
        GRDTranx.TextMatrix(0, 0) = "SL"
        GRDTranx.TextMatrix(0, 1) = "Customer."
        GRDTranx.TextMatrix(0, 2) = "GSTTin No"
        GRDTranx.TextMatrix(0, 3) = "Bill No."
        If Optst.Value = True Then
            GRDTranx.TextMatrix(0, 4) = "Taxable Amt"
        Else
            GRDTranx.TextMatrix(0, 4) = "Bill Amt"
        End If
        GRDTranx.TextMatrix(0, 5) = "Bill Date"
        GRDTranx.ColWidth(0) = 800
        GRDTranx.ColWidth(1) = 3500
        GRDTranx.ColWidth(2) = 1800
        GRDTranx.ColWidth(3) = 1100
        GRDTranx.ColWidth(4) = 1500
        GRDTranx.ColWidth(5) = 1300
        
        GrdTotal.ColWidth(0) = 800
        GrdTotal.ColWidth(1) = 3500
        GrdTotal.ColWidth(2) = 1800
        GrdTotal.ColWidth(3) = 1100
        GrdTotal.ColWidth(4) = 1500
        GrdTotal.ColWidth(5) = 1300
        
        GRDTranx.ColAlignment(0) = 4
        GRDTranx.ColAlignment(1) = 1
        GRDTranx.ColAlignment(2) = 1
        GRDTranx.ColAlignment(3) = 4
        GRDTranx.ColAlignment(4) = 4
        GrdTotal.ColAlignment(3) = 4
        GrdTotal.ColAlignment(4) = 4
        GrdTotal.ColAlignment(5) = 4
        
        n = 6
        M = 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From QTNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='QT' ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Do Until rstTRANX.EOF
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            If OPTGST.Value = True Then
                If rstTRANX!ACT_CODE = "130000" Or rstTRANX!ACT_CODE = "130001" Then
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
                Else
                    GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
                End If
            Else
                GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            End If
            GRDTranx.TextMatrix(M, 2) = IIf(IsNull(rstTRANX!TIN), "", rstTRANX!TIN)
            GRDTranx.TextMatrix(M, 3) = Format(rstTRANX!VCH_NO, bill_for) '& BILL_SUF
            GRDTranx.TextMatrix(M, 5) = IIf(IsDate(rstTRANX!VCH_DATE), rstTRANX!VCH_DATE, "")
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From QTNSUB WHERE TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND VCH_NO = " & rstTRANX!VCH_NO & " ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                GRDTranx.TextMatrix(M, 4) = Round(Val(GRDTranx.TextMatrix(M, 4)) + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL), 3)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing

            DISC_AMT = 0
            DISC_AMT = DISC_AMT + IIf(IsNull(rstTRANX!DISCOUNT), 0, rstTRANX!DISCOUNT)
            GRDTranx.TextMatrix(M, 4) = Format(Round(Val(GRDTranx.TextMatrix(M, 4)) - DISC_AMT, 2), "0.00")
            
            vbalProgressBar1.Value = vbalProgressBar1.Value + 1
            M = M + 1
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        GrdTotal.rows = 0
        GrdTotal.rows = GrdTotal.rows + 1
        GrdTotal.Cols = GRDTranx.Cols
        For n = 4 To GRDTranx.Cols - 1
            If n <> 5 Then
                For i = 1 To GRDTranx.rows - 1
                    GrdTotal.TextMatrix(0, n) = Val(GrdTotal.TextMatrix(0, n)) + Val(GRDTranx.TextMatrix(i, n))
                Next i
            End If
        Next n
    End If
'    GRDTranx.TextMatrix(0, i) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
'    GRDTranx.TextMatrix(0, i + 1) = "Sales " & GRDTranx.TextMatrix(0, i) & "%"
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function

