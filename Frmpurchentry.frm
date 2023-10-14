VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMRECEIPT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   ControlBox      =   0   'False
   Icon            =   "Frmpurchentry.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   9915
   Begin VB.Frame Frame1 
      Height          =   2820
      Left            =   120
      TabIndex        =   0
      Top             =   225
      Width           =   9705
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
         Left            =   1440
         TabIndex        =   21
         Top             =   165
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   1470
         Left            =   5160
         TabIndex        =   12
         Top             =   1200
         Width           =   4485
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal. Amt"
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
            Left            =   2385
            TabIndex        =   20
            Top             =   945
            Width           =   810
         End
         Begin VB.Label Label5 
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   3300
            TabIndex        =   19
            Top             =   885
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Rcvd Amt"
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
            Height          =   465
            Index           =   12
            Left            =   45
            TabIndex        =   18
            Top             =   900
            Width           =   1020
         End
         Begin VB.Label Label4 
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   1200
            TabIndex        =   17
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Dtd. "
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
            Left            =   45
            TabIndex        =   16
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label2 
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   510
            TabIndex        =   15
            Top             =   210
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bill Amt"
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
            Left            =   1845
            TabIndex        =   14
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label Label3 
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
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   3255
            TabIndex        =   13
            Top             =   225
            Width           =   1155
         End
      End
      Begin VB.TextBox txtBillNo 
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
         Left            =   1455
         TabIndex        =   2
         Top             =   1470
         Width           =   885
      End
      Begin VB.TextBox Text1 
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
         Left            =   1455
         TabIndex        =   1
         Top             =   2010
         Width           =   1005
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   3720
         TabIndex        =   3
         Top             =   1470
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmbinv 
         Height          =   330
         Left            =   3960
         TabIndex        =   4
         Top             =   2055
         Visible         =   0   'False
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   255
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
      Begin MSDataListLib.DataList DataList2 
         Height          =   645
         Left            =   1440
         TabIndex        =   22
         Top             =   525
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
      Begin VB.Label Label1 
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
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   23
         Top             =   210
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rcpt. Date"
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
         Left            =   2550
         TabIndex        =   11
         Top             =   1470
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT NO."
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
         Left            =   105
         TabIndex        =   10
         Top             =   1470
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT AMT"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry"
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
         Index           =   3
         Left            =   5835
         TabIndex        =   8
         Top             =   285
         Width           =   1395
      End
      Begin VB.Label LBLDATE 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   7215
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LBLTIME 
         Alignment       =   2  'Center
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
         Height          =   360
         Left            =   8490
         TabIndex        =   6
         Top             =   240
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Against Bill No."
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
         Index           =   4
         Left            =   2520
         TabIndex        =   5
         Top             =   2055
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FRMRECEIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
