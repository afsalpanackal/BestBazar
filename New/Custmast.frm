VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmcustmast1 
   BackColor       =   &H00FAD3EB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Creation"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   Icon            =   "Custmast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   13230
   Begin VB.TextBox txtsupplist 
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
      Left            =   1815
      MaxLength       =   34
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00F7B3DD&
      Height          =   7485
      Left            =   75
      TabIndex        =   16
      Top             =   825
      Width           =   13110
      Begin VB.TextBox txtcrlimit 
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
         Left            =   3195
         MaxLength       =   7
         TabIndex        =   52
         Top             =   3750
         Width           =   1155
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Import Customers"
         Height          =   495
         Left            =   11835
         TabIndex        =   49
         Top             =   390
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Import Customers"
         Height          =   495
         Left            =   11730
         TabIndex        =   48
         Top             =   4050
         Width           =   1305
      End
      Begin VB.CheckBox chkIGST 
         BackColor       =   &H00800080&
         Caption         =   "&IGST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   7275
         TabIndex        =   41
         Top             =   1785
         Width           =   1965
      End
      Begin VB.CheckBox chkdealer 
         BackColor       =   &H00800080&
         Caption         =   "Sub Dealer / Agent"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   7275
         TabIndex        =   40
         Top             =   2145
         Width           =   2640
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00400000&
         Caption         =   "&Add Branch Offices"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   6375
         MaskColor       =   &H80000007&
         TabIndex        =   39
         Top             =   4110
         UseMaskColor    =   -1  'True
         Width           =   1830
      End
      Begin VB.ComboBox cmbtype 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   330
         ItemData        =   "Custmast.frx":0442
         Left            =   7125
         List            =   "Custmast.frx":045B
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   3750
         Width           =   2625
      End
      Begin VB.TextBox Txtopbal 
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
         Left            =   5055
         MaxLength       =   12
         TabIndex        =   34
         Top             =   3750
         Width           =   1470
      End
      Begin VB.TextBox txtcrdtdays 
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
         Left            =   1650
         MaxLength       =   4
         TabIndex        =   32
         Top             =   3750
         Width           =   735
      End
      Begin VB.CheckBox chknewcomp 
         BackColor       =   &H00800080&
         Caption         =   "&New Place"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   4590
         TabIndex        =   29
         Top             =   2565
         Width           =   1920
      End
      Begin VB.TextBox txtcompany 
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
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   1665
         MaxLength       =   20
         TabIndex        =   28
         Top             =   2550
         Width           =   2895
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00400000&
         Caption         =   "&DELETE"
         Enabled         =   0   'False
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
         Left            =   2505
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   6930
         UseMaskColor    =   -1  'True
         Width           =   1170
      End
      Begin VB.TextBox txtcst 
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
         Left            =   4650
         MaxLength       =   25
         TabIndex        =   11
         Top             =   2130
         Width           =   2595
      End
      Begin VB.TextBox txtkgst 
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
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   10
         Top             =   2145
         Width           =   2235
      End
      Begin VB.TextBox txtremarks 
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
         Left            =   915
         MaxLength       =   40
         TabIndex        =   9
         Top             =   6450
         Width           =   4005
      End
      Begin VB.TextBox txtdlno 
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
         Left            =   915
         MaxLength       =   40
         TabIndex        =   8
         Top             =   6075
         Width           =   4005
      End
      Begin VB.TextBox txtemail 
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
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1740
         Width           =   5580
      End
      Begin VB.TextBox txtfaxno 
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
         Left            =   5010
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1350
         Width           =   2235
      End
      Begin VB.TextBox txttelno 
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
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1350
         Width           =   2235
      End
      Begin VB.TextBox txtsupplier 
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
         Left            =   1665
         MaxLength       =   100
         TabIndex        =   3
         Top             =   255
         Width           =   6300
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
         Left            =   60
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   6930
         UseMaskColor    =   -1  'True
         Width           =   1170
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
         Left            =   1275
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   6930
         UseMaskColor    =   -1  'True
         Width           =   1170
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
         Left            =   3735
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   6930
         UseMaskColor    =   -1  'True
         Width           =   1170
      End
      Begin MSDataListLib.DataList Datacompany 
         Height          =   780
         Left            =   1665
         TabIndex        =   30
         Top             =   2925
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1376
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16512
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
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   2835
         Left            =   4920
         TabIndex        =   38
         Top             =   4590
         Width           =   8145
         _ExtentX        =   14367
         _ExtentY        =   5001
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         SelectionMode   =   1
         AllowUserResizing=   1
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DOM 
         Height          =   390
         Left            =   1665
         TabIndex        =   42
         Top             =   4140
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         DateIsNull      =   -1  'True
         Format          =   153485313
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DOB 
         Height          =   390
         Left            =   4785
         TabIndex        =   43
         Top             =   4140
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         CustomFormat    =   "dd/MM"
         DateIsNull      =   -1  'True
         Format          =   153485315
         CurrentDate     =   42543.9362847222
      End
      Begin MSDataListLib.DataCombo CMBDISTI 
         Height          =   1425
         Left            =   1665
         TabIndex        =   46
         Top             =   4590
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   2514
         _Version        =   393216
         Appearance      =   0
         Style           =   1
         ForeColor       =   16711680
         Text            =   ""
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cr. Limit"
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
         Index           =   20
         Left            =   2400
         TabIndex        =   53
         Top             =   3780
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DL NO.1"
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
         Left            =   150
         TabIndex        =   51
         Top             =   6090
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DL NO.2"
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
         Left            =   150
         TabIndex        =   50
         Top             =   6465
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Agent Name"
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
         Left            =   165
         TabIndex        =   47
         Top             =   4635
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Marriage Date"
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
         Left            =   150
         TabIndex        =   45
         Top             =   4200
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
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
         Left            =   3915
         TabIndex        =   44
         Top             =   4200
         Width           =   900
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
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
         Left            =   6555
         TabIndex        =   37
         Top             =   3780
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OP. Bal"
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
         Left            =   4365
         TabIndex        =   35
         Top             =   3780
         Width           =   735
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit days"
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
         Left            =   150
         TabIndex        =   33
         Top             =   3780
         Width           =   1290
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
         Index           =   6
         Left            =   165
         TabIndex        =   31
         Top             =   2715
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UID No."
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
         Left            =   3930
         TabIndex        =   26
         Top             =   2175
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GST No."
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
         Left            =   150
         TabIndex        =   25
         Top             =   2175
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KGST NO.2"
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
         Left            =   9300
         TabIndex        =   24
         Top             =   6240
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KGST NO.1"
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
         Left            =   9255
         TabIndex        =   23
         Top             =   5865
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Left            =   150
         TabIndex        =   22
         Top             =   1785
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
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
         Left            =   3990
         TabIndex        =   21
         Top             =   1395
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
         Left            =   150
         TabIndex        =   20
         Top             =   1395
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
         TabIndex        =   19
         Top             =   645
         Width           =   1290
      End
      Begin MSForms.TextBox txtaddress 
         Height          =   690
         Left            =   1665
         TabIndex        =   4
         Top             =   630
         Width           =   6315
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   99
         BorderStyle     =   1
         Size            =   "11139;1217"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Name"
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
         TabIndex        =   17
         Top             =   255
         Width           =   1515
      End
   End
   Begin VB.TextBox Txtsuplcode 
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
      Left            =   1815
      MaxLength       =   10
      TabIndex        =   0
      Top             =   450
      Width           =   3045
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1500
      Left            =   1815
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2646
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 to Search...... Esc to exit"
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
      Left            =   135
      TabIndex        =   18
      Top             =   45
      Width           =   6300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Code"
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
      Left            =   105
      TabIndex        =   15
      Top             =   465
      Width           =   1560
   End
End
Attribute VB_Name = "frmcustmast1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean

Private Sub Cmdcancel_Click()
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtcrdtdays.Text = ""
    txtcrlimit.Text = ""
    txtremarks.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    CMBDISTI.Text = ""
    txtcompany.Text = ""
    chknewcomp.value = 0
    Txtopbal.Text = ""
    Txtsuplcode.Enabled = True
    chkdealer.value = 0
    chkIGST.value = 0
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From RTRXFILE WHERE M_USER_ID = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRANSMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "SALES") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM CUSTMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "'")
    db.Execute ("delete  FROM PRODLINK WHERE ACT_CODE = '" & Txtsuplcode.Text & "'")
    Call Cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME OF CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If
    If cmbtype.ListIndex = -1 Then
        MsgBox "SELECT TYPE", vbOKOnly, "CUSTOMER MASTER"
        cmbtype.SetFocus
        Exit Sub
    End If
    
'    If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'        MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'        txtcompany.SetFocus
'        Exit Sub
'    End If
    
    If chknewcomp.value = 0 And Datacompany.BoundText = "" And txtcompany.Text <> "" Then
        MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
        txtcompany.SetFocus
        Exit Sub
    End If
    
    If Trim(txtkgst.Text) <> "" Then
        If Len(Trim(txtkgst.Text)) <> 15 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "CUSTOMER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If

        If Val(Left(Trim(txtkgst.Text), 2)) = 0 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "CUSTOMER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If

'        If Val(Mid(Trim(txtkgst.Text), 13, 1)) = 0 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "CUSTOMER MASTER"
'            txtkgst.SetFocus
'            Exit Sub
'        End If

        If Val(Mid(Trim(txtkgst.Text), 14, 1)) <> 0 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "CUSTOMER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If
    End If
    
    On Error GoTo ErrHand
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!ACT_CODE = Txtsuplcode.Text
    End If
    RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
    RSTITEMMAST!Address = Trim(txtaddress.Text)
    RSTITEMMAST!TELNO = Trim(txttelno.Text)
    RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
    RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
    RSTITEMMAST!DL_NO = Trim(txtdlno.Text)
    RSTITEMMAST!REMARKS = Trim(txtremarks.Text)
    RSTITEMMAST!KGST = Trim(txtkgst.Text)
    RSTITEMMAST!CST = Trim(txtcst.Text)
    RSTITEMMAST!PYMT_PERIOD = Val(txtcrdtdays.Text)
    RSTITEMMAST!PYMT_LIMIT = Val(txtcrlimit.Text)
    If txtcompany.Text <> "" Or Datacompany.BoundText <> "" Then
        If chknewcomp.value = 1 Then RSTITEMMAST!Area = txtcompany.Text Else RSTITEMMAST!Area = Datacompany.BoundText
    End If
    If CMBDISTI.BoundText <> "" Then
        RSTITEMMAST!AGENT_CODE = CMBDISTI.BoundText
        RSTITEMMAST!AGENT_NAME = CMBDISTI.Text
    Else
        RSTITEMMAST!AGENT_CODE = ""
        RSTITEMMAST!AGENT_NAME = ""
    End If
    RSTITEMMAST!CONTACT_PERSON = "CS"
    RSTITEMMAST!SLSM_CODE = "SM"
    RSTITEMMAST!OPEN_DB = Round(Val(Txtopbal.Text), 3)
    RSTITEMMAST!OPEN_CR = 0
    RSTITEMMAST!YTD_DB = 0
    RSTITEMMAST!YTD_CR = 0
    RSTITEMMAST!CREATE_DATE = Date
    RSTITEMMAST!C_USER_ID = "SM"
    RSTITEMMAST!MODIFY_DATE = Date
    RSTITEMMAST!M_USER_ID = "SM"
    Select Case cmbtype.ListIndex
        Case 0
            RSTITEMMAST!Type = "R"
        Case 1
            RSTITEMMAST!Type = "W"
        Case 2
            RSTITEMMAST!Type = "V"
        Case 3
            RSTITEMMAST!Type = "M"
        Case 4
            RSTITEMMAST!Type = "5"
        Case 5
            RSTITEMMAST!Type = "6"
        Case 6
            RSTITEMMAST!Type = "7"
        Case Else
            RSTITEMMAST!Type = "R"
    End Select
    RSTITEMMAST!Sl_no = Val(Txtsuplcode.Text)
    If chkdealer.value = 1 Then
        RSTITEMMAST!CUST_TYPE = "D"
    Else
        RSTITEMMAST!CUST_TYPE = ""
    End If
    If chkIGST.value = 1 Then
        RSTITEMMAST!CUST_IGST = "Y"
    Else
        RSTITEMMAST!CUST_IGST = ""
    End If
    RSTITEMMAST!DOM = IIf(IsNull(DOM.value), "", DOM.value)
    RSTITEMMAST!DOB = IIf(IsNull(DOB.value), "", DOB.value)
    
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "CUSTOMER CREATION"
    Dim TRXMAST As ADODB.Recordset
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(SL_NO) From CUSTMAST WHERE (ACT_CODE <> '130000' OR ACT_CODE <> '130001')", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "1", TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select * From CUSTMAST WHERE SL_NO = " & Txtsuplcode.Text & "", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = TRXMAST!ACT_CODE
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtremarks.Text = ""
    txtcrdtdays.Text = ""
    txtcrlimit.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    CMBDISTI.Text = ""
    txtcompany.Text = ""
    Txtopbal.Text = ""
    chknewcomp.value = 0
    chkdealer.value = 0
    chkIGST.value = 0
    DOM.value = Null
    DOB.value = Null
    Txtsuplcode.Enabled = True
    cmdexit.Enabled = True
    cmdcancel.Enabled = True
Exit Sub
ErrHand:
    MsgBox (Err.Description)
        
End Sub

Private Sub Command1_Click()
    Me.Enabled = False
    frmcustTRXFILE.LBLCUSTCODE.Caption = Txtsuplcode.Text
    frmcustTRXFILE.Show
    frmcustTRXFILE.SetFocus
End Sub

Private Sub Command2_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Import Customers"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Import Stock Items") = vbNo Then Exit Sub
    If MsgBox("Sheet Name should be 'CUSTOMERS' and First coloumn should be Customer Name and Second coloumn should be Customer Address", vbYesNo, "Import Customers") = vbNo Then Exit Sub
    On Error GoTo ErrHand
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen

    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application

    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)

    Set ws = wb.Worksheets("CUSTOMERS") 'Specify your worksheet name
    var = ws.Range("A1").value

    'db.Execute "dELETE FROM CUSTMAST"
    'db.Execute "dELETE FROM RTRXFILE"

    Dim RstCustmast As ADODB.Recordset
    Dim RSTITEMTRX As ADODB.Recordset
    Dim CUSTCODE As String
    Dim sl As Integer
    Dim lastno As Integer
    sl = 1

    Set RstCustmast = New ADODB.Recordset
    RstCustmast.Open "Select MAX(SL_NO) From CUSTMAST WHERE (ACT_CODE <> '130000' OR ACT_CODE <> '130001') ", db, adOpenStatic, adLockReadOnly
    If Not (RstCustmast.EOF And RstCustmast.BOF) Then
        If IsNull(RstCustmast.Fields(0)) Then
            CUSTCODE = 1
        Else
            CUSTCODE = Val(RstCustmast.Fields(0)) + 1
        End If
    End If
    RstCustmast.Close
    Set RstCustmast = Nothing

    For i = 2 To 30000
        If Trim(ws.Range("A" & i).value) = "" Then Exit For

        Set RstCustmast = New ADODB.Recordset
        RstCustmast.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & ws.Range("A" & i).value & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans

        If (RstCustmast.EOF And RstCustmast.BOF) Then
            RstCustmast.AddNew
            'RSTCUSTMAST.Fields("PHOTO").AppendChunk bytData
            RstCustmast!ACT_CODE = ws.Range("A" & i).value
            RstCustmast!ACT_NAME = Trim(ws.Range("B" & i).value)
            RstCustmast!Address = Trim(ws.Range("C" & i).value)
            RstCustmast!TELNO = Trim(ws.Range("D" & i).value)
            RstCustmast!FAXNO = Trim(ws.Range("E" & i).value)
            RstCustmast!EMAIL_ADD = ""
            RstCustmast!DL_NO = ""
            RstCustmast!REMARKS = ""
            RstCustmast!KGST = Trim(ws.Range("F" & i).value)
            RstCustmast!CST = ""
            RstCustmast!PYMT_PERIOD = 0
            RstCustmast!Area = Trim(ws.Range("H" & i).value)
            RstCustmast!AGENT_CODE = ""
            RstCustmast!AGENT_NAME = ""
            RstCustmast!Sl_no = ws.Range("A" & i).value
            RstCustmast!CONTACT_PERSON = "CS"
            RstCustmast!SLSM_CODE = "SM"
            RstCustmast!OPEN_DB = ws.Range("G" & i).value
            RstCustmast!OPEN_CR = 0
            RstCustmast!YTD_DB = 0
            RstCustmast!YTD_CR = 0
            RstCustmast!CREATE_DATE = Date
            RstCustmast!C_USER_ID = "SM"
            RstCustmast!MODIFY_DATE = Date
            RstCustmast!M_USER_ID = "SM"
            RstCustmast!Type = "R"
            RstCustmast!CUST_TYPE = ""
            RstCustmast!CUST_IGST = ""

            RstCustmast.Update
            RstCustmast.Close
            Set RstCustmast = Nothing
        End If
        db.CommitTrans
        CUSTCODE = CUSTCODE + 1

SKIP:
    Next i
    wb.Close

    xlApp.Quit

    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbNormal

    MsgBox "Success", vbOKOnly
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = 9 Then
        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf Err.Number = 32755 Then

    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub Command3_Click()
'    Dim rstcustomers As ADODB.Recordset
'    Dim rstcustomers2 As ADODB.Recordset
'    Dim RstCustmast As ADODB.Recordset
'
'    Set rstcustomers = New ADODB.Recordset
'    rstcustomers.Open "SELECT DISTINCT ACT_CODE FROM dbtpymt", db, adOpenStatic, adLockReadOnly, adCmdText
'
'    Set RstCustmast = New ADODB.Recordset
'    RstCustmast.Open "SELECT * FROM CUSTMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until rstcustomers.EOF
'        RstCustmast.AddNew
'        'RSTCUSTMAST.Fields("PHOTO").AppendChunk bytData
'        RstCustmast!act_code = rstcustomers!act_code
'        Set rstcustomers2 = New ADODB.Recordset
'        rstcustomers2.Open "SELECT * from dbtpymt where ACT_CODE = '" & rstcustomers!act_code & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rstcustomers2.EOF And rstcustomers2.BOF) Then
'            RstCustmast!act_name = rstcustomers2!act_name
'        End If
'        rstcustomers2.Close
'        Set rstcustomers2 = Nothing
'
'        RstCustmast!Address = ""
'        RstCustmast!TELNO = ""
'        RstCustmast!FAXNO = ""
'        RstCustmast!EMAIL_ADD = ""
'        RstCustmast!DL_NO = ""
'        RstCustmast!Remarks = ""
'        RstCustmast!KGST = ""
'        RstCustmast!CST = ""
'        RstCustmast!PYMT_PERIOD = 0
'        RstCustmast!Area = ""
'        RstCustmast!AGENT_CODE = ""
'        RstCustmast!AGENT_NAME = ""
'        RstCustmast!Sl_no = CUSTCODE
'        RstCustmast!CONTACT_PERSON = "CS"
'        RstCustmast!SLSM_CODE = "SM"
'        RstCustmast!OPEN_DB = 0
'        RstCustmast!OPEN_CR = 0
'        RstCustmast!YTD_DB = 0
'        RstCustmast!YTD_CR = 0
'        RstCustmast!CREATE_DATE = Date
'        RstCustmast!C_USER_ID = "SM"
'        RstCustmast!MODIFY_DATE = Date
'        RstCustmast!M_USER_ID = "SM"
'        RstCustmast!Type = "W"
'        RstCustmast!CUST_TYPE = ""
'        RstCustmast!CUST_IGST = ""
'        RstCustmast.Update
'
'        rstcustomers.MoveNext
'    Loop
'    RstCustmast.Close
'    Set RstCustmast = Nothing
'
'    rstcustomers.Close
'    Set rstcustomers = Nothing
'
'    Set RstCustmast = New ADODB.Recordset
'    RstCustmast.Open "SELECT * FROM trxsub", db, adOpenStatic, adLockOptimistic, adCmdText
'
'    Set rstcustomers = New ADODB.Recordset
'    rstcustomers.Open "SELECT * FROM trxfile", db, adOpenStatic, adLockReadOnly, adCmdText
'    Do Until rstcustomers.EOF
'        RstCustmast.AddNew
'        RstCustmast!VCH_NO = rstcustomers!VCH_NO
'        RstCustmast!TRX_TYPE = rstcustomers!TRX_TYPE
'        RstCustmast!LINE_NO = rstcustomers!LINE_NO
'        RstCustmast!TRX_YEAR = rstcustomers!TRX_YEAR
'        RstCustmast!R_VCH_NO = 1
'        RstCustmast!R_TRX_TYPE = "OP"
'        RstCustmast!R_LINE_NO = 1
'        RstCustmast!R_TRX_YEAR = "2018"
'        RstCustmast!QTY = rstcustomers!QTY
'        RstCustmast.Update
'        rstcustomers.MoveNext
'    Loop
'    RstCustmast.Close
'    Set RstCustmast = Nothing
'    Exit Sub
'Errhand:
'    MsgBox Err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ACT_CODE FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.Text = RSTITEMMAST!ACT_CODE
            End If
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    'If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    AGNT_FLAG = True
    'TMPFLAG = True
    'Me.Width = 7000
    'Me.Height = 8625
    Me.Left = 2500
    Me.Top = 0
    FRAME.Visible = False
    'txtunit.Visible = False
    On Error GoTo ErrHand
    
    Call FILLCOMBO
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(SL_NO) From CUSTMAST WHERE (ACT_CODE <> '130000' OR ACT_CODE <> '130001')", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "1", TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select * From CUSTMAST WHERE SL_NO = " & Txtsuplcode.Text & "", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = TRXMAST!ACT_CODE
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing

    grdsales.TextMatrix(0, 0) = "SL"
    grdsales.TextMatrix(0, 2) = "Branch Name"
    grdsales.TextMatrix(0, 3) = "Address"
    
    grdsales.ColWidth(0) = 500
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2800
    grdsales.ColWidth(3) = 5000
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(2) = 4
    grdsales.ColAlignment(3) = 4
    DOM.value = Null
    DOB.value = Null
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    If AGNT_FLAG = False Then ACT_AGNT.Close
    
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub txtremarks_GotFocus()
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
        Case vbKeyEscape
            txtdlno.SetFocus
    End Select
End Sub

Private Sub txtaddress_GotFocus()
    txtaddress.SelStart = 0
    txtaddress.SelLength = Len(txtaddress.Text)
End Sub

Private Sub txtaddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'txttelno.SetFocus
        Case vbKeyEscape
            txtsupplier.SetFocus
    End Select
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtcst_GotFocus()
    txtcst.SelStart = 0
    txtcst.SelLength = Len(txtcst.Text)
End Sub

Private Sub txtcst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcompany.SetFocus
        Case vbKeyEscape
            txtkgst.SetFocus
    End Select
End Sub

Private Sub txtdlno_GotFocus()
    txtdlno.SelStart = 0
    txtdlno.SelLength = Len(txtdlno.Text)
End Sub

Private Sub txtdlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtcompany.SetFocus
        Case vbKeyReturn
            txtremarks.SetFocus
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
        Case vbKeyEscape
            txtfaxno.SetFocus
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
        Case vbKeyEscape
            txttelno.SetFocus
    End Select
End Sub

Private Sub txtkgst_GotFocus()
    txtkgst.SelStart = 0
    txtkgst.SelLength = Len(txtkgst.Text)
End Sub

Private Sub txtkgst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcst.SetFocus
        Case vbKeyEscape
            txtemail.SetFocus
    End Select
End Sub

Private Sub txtkgst_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z")
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
                MsgBox "ENTER NAME OF CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
                txtsupplier.SetFocus
                Exit Sub
            End If
         txtaddress.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub Txtsuplcode_GotFocus()
    Txtsuplcode.SelStart = 0
    Txtsuplcode.SelLength = Len(Txtsuplcode.Text)
End Sub

Private Sub Txtsuplcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Txtsuplcode.Text) = "" Then Exit Sub
            'If Val(Txtsuplcode.Text) = 0 Then Exit Sub
            If Trim(Txtsuplcode.Text) = "130000" Or Trim(Txtsuplcode.Text) = "130001" Then
                MsgBox "This Code Cannot be created!!!!", , "Customer Creation"
                Exit Sub
            End If
            
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            'RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "' and ACT_CODE <> '130000'", db, adOpenStatic, adLockReadOnly
            RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & Txtsuplcode.Text & "' ", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.Text = RSTITEMMAST!ACT_NAME
                txtaddress.Text = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                txttelno.Text = IIf(IsNull(RSTITEMMAST!TELNO), "", RSTITEMMAST!TELNO)
                txtfaxno.Text = IIf(IsNull(RSTITEMMAST!FAXNO), "", RSTITEMMAST!FAXNO)
                txtemail.Text = IIf(IsNull(RSTITEMMAST!EMAIL_ADD), "", RSTITEMMAST!EMAIL_ADD)
                txtdlno.Text = IIf(IsNull(RSTITEMMAST!DL_NO), "", RSTITEMMAST!DL_NO)
                txtremarks.Text = IIf(IsNull(RSTITEMMAST!REMARKS), "", RSTITEMMAST!REMARKS)
                txtkgst.Text = IIf(IsNull(RSTITEMMAST!KGST), "", RSTITEMMAST!KGST)
                txtcst.Text = IIf(IsNull(RSTITEMMAST!CST), "", RSTITEMMAST!CST)
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!Area), "", RSTITEMMAST!Area)
                CMBDISTI.Text = IIf(IsNull(RSTITEMMAST!AGENT_NAME), "", RSTITEMMAST!AGENT_NAME)
                CMBDISTI.BoundText = IIf(IsNull(RSTITEMMAST!AGENT_CODE), "", RSTITEMMAST!AGENT_CODE)
                Txtopbal.Text = IIf(IsNull(RSTITEMMAST!OPEN_DB), 0, RSTITEMMAST!OPEN_DB)
                txtcrdtdays.Text = IIf(IsNull(RSTITEMMAST!PYMT_PERIOD), 0, RSTITEMMAST!PYMT_PERIOD)
                txtcrlimit.Text = IIf(IsNull(RSTITEMMAST!PYMT_LIMIT), 0, RSTITEMMAST!PYMT_LIMIT)
                Select Case RSTITEMMAST!Type
                    Case "W"
                        cmbtype.ListIndex = 1
                    Case "V"
                        cmbtype.ListIndex = 2
                    Case "M"
                        cmbtype.ListIndex = 3
                    Case "5"
                        cmbtype.ListIndex = 4
                    Case "6"
                        cmbtype.ListIndex = 5
                    Case "7"
                        cmbtype.ListIndex = 6
                    Case Else
                        cmbtype.ListIndex = 0
                End Select
                
                DOM.value = IIf(IsDate(RSTITEMMAST!DOM), Format(RSTITEMMAST!DOM, "dd/MM/yyyy"), Null)
                DOB.value = IIf(IsDate(RSTITEMMAST!DOB), Format(RSTITEMMAST!DOB, "DD/MM"), Null)
                If RSTITEMMAST!CUST_TYPE = "D" Then
                    chkdealer.value = 1
                Else
                    chkdealer.value = 0
                End If
                If RSTITEMMAST!CUST_IGST = "Y" Then
                    chkIGST.value = 1
                Else
                    chkIGST.value = 0
                End If
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
                cmddelete.Enabled = True
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Dim i As Long
            i = 1
            grdsales.FixedRows = 0
            grdsales.Rows = 1
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & Txtsuplcode.Text & "' ", db, adOpenStatic, adLockReadOnly
            Do Until RSTITEMMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = RSTITEMMAST!BR_CODE
                grdsales.TextMatrix(i, 2) = IIf(IsNull(RSTITEMMAST!BR_NAME), "", RSTITEMMAST!BR_NAME)
                grdsales.TextMatrix(i, 3) = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                RSTITEMMAST.MoveNext
                i = i + 1
                grdsales.FixedRows = 1
            Loop
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Txtsuplcode.Enabled = False
            FRAME.Visible = True
            txtsupplier.SetFocus
        Case 114
            txtsupplist.Text = ""
            txtsupplist.Visible = True
            DataList2.Visible = True
            txtsupplist.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Txtsuplcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplist_Change()
    On Error GoTo ErrHand
    If REPFLAG = True Then
        RSTREP.Open "Select ACT_CODE,ACT_NAME From CUSTMAST  WHERE (ACT_CODE <> '130000' OR ACT_CODE <> '130001') And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ACT_CODE,ACT_NAME From CUSTMAST  WHERE (ACT_CODE <> '130000' OR ACT_CODE <> '130001') And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        'RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic,adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ACT_NAME"
    DataList2.BoundColumn = "ACT_CODE"
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub txtsupplist_GotFocus()
    txtsupplist.SelStart = 0
    txtsupplist.SelLength = Len(txtsupplist.Text)
End Sub

Private Sub txtsupplist_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtsupplist.Text = "" Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub txtsupplist_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txttelno_GotFocus()
    txttelno.SelStart = 0
    txttelno.SelLength = Len(txttelno.Text)
End Sub

Private Sub txttelno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtfaxno.SetFocus
        Case vbKeyEscape
            txtaddress.SetFocus
    End Select
End Sub


Private Sub txtcompany_Change()
    On Error GoTo ErrHand
    
    Set Me.Datacompany.RowSource = Nothing
    If COMPANYFLAG = True Then
        RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST WHERE AREA Like '" & txtcompany.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    Else
        RSTCOMPANY.Close
        RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST WHERE AREA Like '" & txtcompany.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    End If
    Set Me.Datacompany.RowSource = RSTCOMPANY
    Datacompany.ListField = "AREA"
    Datacompany.BoundColumn = "AREA"
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub txtcompany_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.Text)
End Sub

Private Sub txtcompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            '''''''If txtcompany.Text = "" Then Exit Sub
            Datacompany.SetFocus
        Case vbKeyEscape
            txtcst.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT AREA FROM CUSTMAST WHERE ACT_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtcompany.Text = RSTITEMMAST!Area
            Else
'                If txtcompany.Text = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
'                If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
                If chknewcomp.value = 0 And Datacompany.BoundText = "" And txtcompany.Text <> "" Then
                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
                    txtcompany.SetFocus
                    Exit Sub
                End If
            End If
            txtcrdtdays.SetFocus
        Case vbKeyEscape
            txtcompany.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub Datacompany_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ErrHand
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEM_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        txtcompany.Text = RSTITEMMAST!MANUFACTURER
'    End If
    txtcompany.Text = Datacompany.BoundText
    Datacompany.Text = txtcompany.Text
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub chknewcomp_Click()
    On Error Resume Next
    txtcompany.SetFocus
End Sub

Private Sub txtcrdtdays_GotFocus()
    txtcrdtdays.SelStart = 0
    txtcrdtdays.SelLength = Len(txtcrdtdays.Text)
End Sub

Private Sub txtcrdtdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcrlimit.SetFocus
        Case vbKeyEscape
            Datacompany.SetFocus
    End Select
End Sub

Private Sub txtcrdtdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtopbal_GotFocus()
    Txtopbal.SelStart = 0
    Txtopbal.SelLength = Len(Txtopbal.Text)
End Sub

Private Sub Txtopbal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmbtype.SetFocus
        Case vbKeyEscape
            txtcrlimit.SetFocus
    End Select
End Sub

Private Sub Txtopbal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("."), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmbtype_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If cmbtype.ListIndex = -1 Then
                MsgBox "SELECT TYPE", vbOKOnly, "CUSTOMER MASTER"
                cmbtype.SetFocus
                Exit Sub
            End If
            CMBDISTI.SetFocus
        Case vbKeyEscape
            Txtopbal.SetFocus
    End Select
End Sub

Private Function FILLCOMBO()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='911')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If CMBDISTI.Text = "" Then Exit Sub
            If IsNull(CMBDISTI.SelectedItem) And CMBDISTI.Text <> "" Then
                MsgBox "Select Agent From List", vbOKOnly, "Customer Creation"
                CMBDISTI.SetFocus
                Exit Sub
            End If
            
'            If Trim(TXTAREA.Text) = "" Then
'                MsgBox "Enter Area for the Customer", vbOKOnly, "Sale Bill..."
'                TXTAREA.SetFocus
'                Exit Sub
'            End If
            
'            If Not IsDate(TXTINVDATE.Text) Then
'                MsgBox "Enter Proper date for Invoice", vbOKOnly, "Sale Bill..."
'                TXTINVDATE.SetFocus
'                Exit Sub
'            End If
'
            'FRMEHEAD.Enabled = False
            CmdSave.SetFocus
        Case vbKeyEscape
            cmbtype.Enabled = True
            cmbtype.SetFocus
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtcrlimit_GotFocus()
    txtcrlimit.SelStart = 0
    txtcrlimit.SelLength = Len(txtcrlimit.Text)
End Sub

Private Sub txtcrlimit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtopbal.SetFocus
        Case vbKeyEscape
            txtcrdtdays.SetFocus
    End Select
End Sub

Private Sub txtcrlimit_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
