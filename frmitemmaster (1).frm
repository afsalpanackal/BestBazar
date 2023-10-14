VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmitemmastermini 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Item Creation"
   ClientHeight    =   5010
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   8070
   ControlBox      =   0   'False
   Icon            =   "frmitemmaster.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   8070
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   9285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   7110
      ScaleHeight     =   300
      ScaleWidth      =   750
      TabIndex        =   22
      Top             =   120
      Width           =   750
   End
   Begin VB.TextBox TxtItemcode 
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
      Height          =   435
      Left            =   3780
      MaxLength       =   21
      TabIndex        =   17
      Top             =   4740
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   -45
      Width           =   8025
      Begin VB.TextBox txtpackdet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   7050
         MaxLength       =   5
         TabIndex        =   47
         Top             =   8160
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.TextBox Txtpackdes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   5865
         MaxLength       =   5
         TabIndex        =   45
         Top             =   8160
         Visible         =   0   'False
         Width           =   570
      End
      Begin VB.TextBox Txtbarcode 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   43
         Top             =   1290
         Width           =   2985
      End
      Begin VB.TextBox TxtMalay 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Kerala"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   41
         Top             =   825
         Width           =   4740
      End
      Begin VB.TextBox TxtMRP 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   1710
         TabIndex        =   39
         Top             =   1755
         Width           =   855
      End
      Begin VB.TextBox TxtLocation 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   6465
         MaxLength       =   10
         TabIndex        =   37
         Top             =   4395
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.TextBox TxtHSN 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   3285
         MaxLength       =   10
         TabIndex        =   35
         Top             =   2655
         Width           =   1410
      End
      Begin VB.CheckBox cHKcHANGE 
         Caption         =   "Price Changing Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   1725
         TabIndex        =   34
         Top             =   4080
         Width           =   2160
      End
      Begin VB.CheckBox chkunbill 
         Caption         =   "Un Bill Items"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   270
         Left            =   60
         TabIndex        =   33
         Top             =   4065
         Width           =   1575
      End
      Begin VB.TextBox txtPack 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   7605
         TabIndex        =   3
         Top             =   8295
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.TextBox txtLPrice 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   2280
         TabIndex        =   11
         Top             =   4395
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.TextBox TxtCost 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   3540
         TabIndex        =   7
         Top             =   2205
         Width           =   1155
      End
      Begin VB.TextBox TxtLPack 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   975
         TabIndex        =   10
         Top             =   4395
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox TxtWS 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   6990
         TabIndex        =   9
         Top             =   3930
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.TextBox txtRT 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   3720
         TabIndex        =   8
         Top             =   1755
         Width           =   975
      End
      Begin VB.TextBox TxtTax 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   1710
         TabIndex        =   6
         Top             =   2205
         Width           =   555
      End
      Begin VB.ComboBox cmbfullpack 
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
         ItemData        =   "frmitemmaster.frx":000C
         Left            =   5085
         List            =   "frmitemmaster.frx":0052
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   8190
         Visible         =   0   'False
         Width           =   1020
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
         ItemData        =   "frmitemmaster.frx":00D1
         Left            =   1710
         List            =   "frmitemmaster.frx":0117
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2685
         Width           =   1005
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
         Left            =   6690
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   3330
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Left            =   5370
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   3330
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox TxtMinQty 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   420
         Left            =   6570
         TabIndex        =   2
         Top             =   8295
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.CommandButton CMDDELETE 
         BackColor       =   &H00400000&
         Caption         =   "&Delete"
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
         Left            =   6675
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   3840
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox TXTITEM 
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
         Height          =   450
         Left            =   1710
         TabIndex        =   1
         Top             =   345
         Width           =   5910
      End
      Begin VB.Label LBLPACKDESC 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack Details"
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
         Height          =   435
         Index           =   0
         Left            =   6450
         TabIndex        =   48
         Top             =   8160
         Visible         =   0   'False
         Width           =   1020
      End
      Begin VB.Label LBLPACKDESC 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "No. of Items in Box"
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
         Height          =   435
         Index           =   20
         Left            =   4830
         TabIndex        =   46
         Top             =   8160
         Visible         =   0   'False
         Width           =   1020
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
         Height          =   240
         Index           =   19
         Left            =   90
         TabIndex        =   44
         Top             =   1380
         Width           =   1635
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Malayalam"
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
         Height          =   240
         Index           =   18
         Left            =   90
         TabIndex        =   42
         Top             =   900
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MRP"
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
         Left            =   135
         TabIndex        =   40
         Top             =   1830
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Loc"
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
         Left            =   6075
         TabIndex        =   38
         Top             =   4445
         Visible         =   0   'False
         Width           =   525
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   14
         Left            =   2805
         TabIndex        =   36
         Top             =   2715
         Width           =   450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "L. Price"
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
         Left            =   1560
         TabIndex        =   32
         Top             =   4440
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label lblPack 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2130
         TabIndex        =   31
         Top             =   5055
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Tax"
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
         Left            =   105
         TabIndex        =   30
         Top             =   2280
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "L. Pack"
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
         Left            =   60
         TabIndex        =   29
         Top             =   4440
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "W. Price"
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
         Left            =   6180
         TabIndex        =   28
         Top             =   3990
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "R. Price"
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
         Left            =   2685
         TabIndex        =   27
         Top             =   1815
         Width           =   780
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cost"
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
         Left            =   3000
         TabIndex        =   26
         Top             =   2295
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "L. Pack"
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
         Left            =   4410
         TabIndex        =   24
         Top             =   8220
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Pack"
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
         TabIndex        =   23
         Top             =   2760
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-order"
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
         Left            =   5640
         TabIndex        =   16
         Top             =   8355
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
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
         Left            =   75
         TabIndex        =   15
         Top             =   420
         Width           =   1995
      End
   End
   Begin VB.Label LBLITEMNAME 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   8910
      TabIndex        =   25
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   6600
      TabIndex        =   21
      Top             =   375
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   7680
      TabIndex        =   20
      Top             =   210
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   7290
      TabIndex        =   19
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   6420
      TabIndex        =   18
      Top             =   60
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmitemmastermini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset

Private Sub cmbfullpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            If cmbfullpack.ListIndex = -1 Then cmbfullpack.Text = CmbPack.Text
            Txtpackdes.SetFocus
        Case vbKeyEscape
            CmbPack.SetFocus
    End Select
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            txtHSN.SetFocus
        Case vbKeyEscape
            TxtCost.SetFocus
    End Select
End Sub

Private Sub CmbPack_LostFocus()
    LblPack.Caption = CmbPack.Text
End Sub


Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim wid As Single
    Dim hgt As Single
    Dim i As Long
    
    On Error GoTo Errhand
    
    i = Val(InputBox("Enter number of lables to be print", "No. of labels.."))
    If i = 0 Then i = 1
    Do Until i = 0
        Picture2.ScaleMode = vbPixels
        Printer.PaintPicture Picture2.Image, 500, 500 ', wid, hgt
        
    '   wid = ScaleX(Picture1.ScaleWidth, Picture1.ScaleMode, _
    '    Printer.ScaleMode)
    '    hgt = ScaleY(Picture1.ScaleHeight, _
    '    Picture1.ScaleMode, Printer.ScaleMode)
    '
    '    ' Draw the box.
    '    Printer.Line (1440, 1440)-Step(wid, hgt), , B
        
        ' Finish printing.
        Printer.EndDoc
        i = i - 1
    Loop
    MsgBox "Done"
    
    Exit Sub
Errhand:
    MsgBox Err.Description

End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    On Error Resume Next
    If Val(Txtpackdes.Text) = 0 Then Txtpackdes.Text = 1
    If Val(txtpackdet.Text) = 0 Then txtpackdet.Text = 1
    If TXTITEM.Text = "" Then
        MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    If Val(txtRT.Text) = 0 Then
        MsgBox "Please enter the Selling Price", vbOKOnly, "PRODUCT MASTER"
        txtRT.SetFocus
        Exit Sub
    End If
    On Error GoTo Errhand
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME = '" & Trim(TXTITEM.Text) & "' AND ITEM_CODE <> '" & Trim(TxtItemcode.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        MsgBox "The Item name already exists...", vbOKOnly, "Item Master"
'        RSTITEMMAST.Close
'        Set RSTITEMMAST = Nothing
'        Exit Sub
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        If IsNull(RSTITEMMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(RSTITEMMAST.Fields(0)) + 1
        End If
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        RSTITEMMAST!Category = "GENERAL"
        RSTITEMMAST!UNIT = 1
        RSTITEMMAST!MANUFACTURER = "GENERAL"
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        RSTITEMMAST!PACK_TYPE = CmbPack.Text
        RSTITEMMAST!FULL_PACK = cmbfullpack.Text
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.Text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.Text)
        RSTITEMMAST!SALES_TAX = Val(TXTTAX.Text)
        RSTITEMMAST!Remarks = Trim(txtHSN.Text)
        RSTITEMMAST!ITEM_SPEC = ""
        If Val(TXTTAX.Text) > 0 Then RSTITEMMAST!CHECK_FLAG = "V"
        RSTITEMMAST!ITEM_COST = Val(TxtCost.Text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.Text)
        RSTITEMMAST!MRP = Val(TxtMRP.Text)
        RSTITEMMAST!P_WS = Val(txtWS.Text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.Text) = 0, 1, Val(TxtLPack.Text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.Text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.Text) = 0, 1, Val(Txtpack.Text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.Text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.Text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.Text)
        RSTITEMMAST!DEAD_STOCK = "N"
        If chkunbill.value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        If cHKcHANGE.value = 0 Then
            RSTITEMMAST!PRICE_CHANGE = "N"
        Else
            RSTITEMMAST!PRICE_CHANGE = "Y"
        End If
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!ITEM_CODE = Trim(TXTITEMCODE.Text)
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        RSTITEMMAST!Category = "GENERAL"
        RSTITEMMAST!MANUFACTURER = "GENERAL"
        RSTITEMMAST!UNIT = 1
        RSTITEMMAST!DEAD_STOCK = "N"
        If chkunbill.value = 0 Then
            RSTITEMMAST!UN_BILL = "N"
        Else
            RSTITEMMAST!UN_BILL = "Y"
        End If
        If cHKcHANGE.value = 0 Then
            RSTITEMMAST!PRICE_CHANGE = "N"
        Else
            RSTITEMMAST!PRICE_CHANGE = "Y"
        End If
        
        RSTITEMMAST!Remarks = Trim(txtHSN.Text)
        RSTITEMMAST!ITEM_SPEC = ""
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        If CmbPack.ListIndex = -1 Then
            RSTITEMMAST!PACK_TYPE = "Nos"
        Else
            RSTITEMMAST!PACK_TYPE = CmbPack.Text
        End If
        If cmbfullpack.ListIndex = -1 Then
            RSTITEMMAST!FULL_PACK = "Nos"
        Else
            RSTITEMMAST!FULL_PACK = CmbPack.Text
        End If
        RSTITEMMAST!BIN_LOCATION = Trim(TxtLocation.Text)
        RSTITEMMAST!ITEM_MAL = Trim(TxtMalay.Text)
        RSTITEMMAST!PTR = 0
        RSTITEMMAST!CST = 0
        RSTITEMMAST!OPEN_QTY = 0
        RSTITEMMAST!OPEN_VAL = 0
        RSTITEMMAST!RCPT_QTY = 0
        RSTITEMMAST!RCPT_VAL = 0
        RSTITEMMAST!ISSUE_QTY = 0
        RSTITEMMAST!ISSUE_VAL = 0
        RSTITEMMAST!CLOSE_QTY = 0
        RSTITEMMAST!CLOSE_VAL = 0
        RSTITEMMAST!DAM_QTY = 0
        RSTITEMMAST!DAM_VAL = 0
        RSTITEMMAST!DISC = 0
        RSTITEMMAST!SALES_TAX = Val(TXTTAX.Text)
        If Val(TXTTAX.Text) > 0 Then RSTITEMMAST!CHECK_FLAG = "V"
        RSTITEMMAST!ITEM_COST = Val(TxtCost.Text)
        RSTITEMMAST!P_RETAIL = Val(txtRT.Text)
        RSTITEMMAST!MRP = Val(TxtMRP.Text)
        RSTITEMMAST!P_WS = Val(txtWS.Text)
        RSTITEMMAST!CRTN_PACK = IIf(Val(TxtLPack.Text) = 0, 1, Val(TxtLPack.Text))
        RSTITEMMAST!P_CRTN = Val(txtLPrice.Text)
        RSTITEMMAST!LOOSE_PACK = IIf(Val(Txtpack.Text) = 0, 1, Val(Txtpack.Text))
        RSTITEMMAST!PACK_DESC = Val(Txtpackdes.Text)
        RSTITEMMAST!PACK_DET = Val(txtpackdet.Text)
        RSTITEMMAST!BARCODE = Trim(TxtBarcode.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTITEMMAST.EOF
'        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
'        RSTITEMMAST!MFGR = Trim(txtcompany.Text)
'        RSTITEMMAST!Category = Trim(txtcategory.Text)
'        RSTITEMMAST.Update
'        RSTITEMMAST.MoveNext
'    Loop
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    
                    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    Unload Me
Exit Sub
Errhand:
    MsgBox (Err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtHSN.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    
    REPFLAG = True
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo Errhand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ITEMMAST ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    If PC_FLAG = "Y" Then
        cHKcHANGE.value = 1
    Else
        cHKcHANGE.value = 0
    End If
    
    'TMPFLAG = True
    'Width = 8385
    'Height = 4575
    Left = 3500
    Top = 0
    FRAME.Visible = False
    'txtunit.Visible = False
    
    Picture2.ScaleMode = 3
    Picture2.Height = Picture2.Height * (1.4 * 40 / Picture2.ScaleHeight)
    Picture2.FontSize = 8

    Exit Sub
Errhand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    'If TMPFLAG = False Then rstTMP.Close
    'MDIMAIN.Enabled = True
    'FrmCrimedata.Enabled = True
End Sub

Private Sub TxtCost_GotFocus()
    TxtCost.SelStart = 0
    TxtCost.SelLength = Len(TxtCost.Text)
End Sub

Private Sub TxtCost_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.SetFocus
        Case vbKeyEscape
            TXTTAX.SetFocus
    End Select
End Sub

Private Sub TxtCost_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtCost_LostFocus()
    TxtCost.Text = Format(Val(TxtCost.Text), "0.00")
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then
                MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            TxtMalay.SetFocus
    End Select
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
End Sub

Public Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo Errhand
            If Trim(TXTITEMCODE.Text) = "" Then Exit Sub
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTITEM.Text = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
                TXTTAX.Text = IIf(IsNull(RSTITEMMAST!SALES_TAX), 0, RSTITEMMAST!SALES_TAX)
                txtHSN.Text = IIf(IsNull(RSTITEMMAST!Remarks), "", RSTITEMMAST!Remarks)
                TxtBarcode.Text = IIf(IsNull(RSTITEMMAST!BARCODE), "", RSTITEMMAST!BARCODE)
                TxtLocation.Text = IIf(IsNull(RSTITEMMAST!BIN_LOCATION), "", RSTITEMMAST!BIN_LOCATION)
                TxtMalay.Text = IIf(IsNull(RSTITEMMAST!ITEM_MAL), "", RSTITEMMAST!ITEM_MAL)
                TxtCost.Text = IIf(IsNull(RSTITEMMAST!ITEM_COST), 0, RSTITEMMAST!ITEM_COST)
                txtRT.Text = IIf(IsNull(RSTITEMMAST!P_RETAIL), 0, RSTITEMMAST!P_RETAIL)
                txtWS.Text = IIf(IsNull(RSTITEMMAST!P_WS), 0, RSTITEMMAST!P_WS)
                TxtMRP.Text = IIf(IsNull(RSTITEMMAST!MRP), 0, RSTITEMMAST!MRP)
                TxtLPack.Text = IIf(IsNull(RSTITEMMAST!CRTN_PACK), 1, RSTITEMMAST!CRTN_PACK)
                txtLPrice.Text = IIf(IsNull(RSTITEMMAST!P_CRTN), 0, RSTITEMMAST!P_CRTN)
                Txtpackdes.Text = IIf(IsNull(RSTITEMMAST!PACK_DESC), 1, RSTITEMMAST!PACK_DESC)
                txtpackdet.Text = IIf(IsNull(RSTITEMMAST!PACK_DET), 1, RSTITEMMAST!PACK_DET)
                Txtpack.Text = IIf(IsNull(RSTITEMMAST!LOOSE_PACK), 1, RSTITEMMAST!LOOSE_PACK)
                If IsNull(RSTITEMMAST!UN_BILL) Or RSTITEMMAST!UN_BILL = "N" Or RSTITEMMAST!UN_BILL = "" Then
                    chkunbill.value = 0
                Else
                    chkunbill.value = 1
                End If
                If IsNull(RSTITEMMAST!PRICE_CHANGE) Or RSTITEMMAST!PRICE_CHANGE = "N" Or RSTITEMMAST!PRICE_CHANGE = "" Then
                    cHKcHANGE.value = 0
                Else
                    cHKcHANGE.value = 1
                End If
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTITEMMAST!PACK_TYPE), 0, RSTITEMMAST!PACK_TYPE)
                cmbfullpack.Text = IIf(IsNull(RSTITEMMAST!FULL_PACK), 0, RSTITEMMAST!FULL_PACK)
                If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
                If Val(TxtLPack.Text) = 0 Then TxtLPack.Text = 1
                If Val(txtLPrice.Text) = 0 Then txtLPrice.Text = Val(txtRT.Text) / Val(TxtLPack.Text)
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        
            TXTITEMCODE.Enabled = False
            FRAME.Visible = True
            TXTITEM.SetFocus
        Case vbKeyEscape
            Call CMDEXIT_Click
    End Select
Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLPack_Change()
    txtLPrice.Text = ""
End Sub

Private Sub TxtLPack_GotFocus()
    If Val(TxtLPack.Text) = 0 Then
        TxtLPack.Text = 1
    End If
    TxtLPack.SelStart = 0
    TxtLPack.SelLength = Len(TxtLPack.Text)
End Sub

Private Sub TxtLPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtLPrice.SetFocus
        Case vbKeyEscape
            txtWS.SetFocus
    End Select
End Sub

Private Sub TxtLPack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLPrice_GotFocus()
    If Val(TxtLPack.Text) = 0 Then TxtLPack.Text = "1"
    If Val(txtLPrice.Text) = 0 Then
        txtLPrice.Text = Format(Round(Val(txtRT.Text) / Val(TxtLPack.Text), 2), "0.00")
    End If
    txtLPrice.SelStart = 0
    txtLPrice.SelLength = Len(txtLPrice.Text)
End Sub

Private Sub txtLPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtHSN.SetFocus
        Case vbKeyEscape
            TxtLPack.SetFocus
    End Select
End Sub

Private Sub txtLPrice_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLPrice_LostFocus()
    txtLPrice.Text = Format(Val(txtLPrice.Text), "0.00")
End Sub

Private Sub txtminqty_GotFocus()
    If TxtMinQty.Text = "" Then
        TxtMinQty.Text = 1
    End If
    TxtMinQty.SelStart = 0
    TxtMinQty.SelLength = Len(TxtMinQty.Text)
End Sub

Private Sub Txtpack_GotFocus()
    If Val(Txtpack.Text) = 0 Then
        Txtpack.Text = 1
    End If
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.Text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmbPack.SetFocus
        Case vbKeyEscape
            TxtMinQty.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub CmdDelete_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    
    If TXTITEMCODE.Text = "" Then Exit Sub
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available in Formula", vbCritical, "DELETING ITEM...."
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & TXTITEM.Text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
    MsgBox "ITEM " & TXTITEM.Text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    Unload Me
    Exit Sub
   
Errhand:
    MsgBox Err.Description
End Sub

Private Sub txtRT_GotFocus()
    txtRT.SelStart = 0
    txtRT.SelLength = Len(txtRT.Text)
End Sub

Private Sub txtRT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtRT.Text) = 0 Then Exit Sub
            TXTTAX.SetFocus
        Case vbKeyEscape
            TxtMRP.SetFocus
    End Select
End Sub

Private Sub txtRT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtRT_LostFocus()
    txtRT.Text = Format(Val(txtRT.Text), "0.00")
End Sub

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCost.SetFocus
        Case vbKeyEscape
            txtRT.SetFocus
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

Private Sub TXTTAX_LostFocus()
    TXTTAX.Text = Format(Val(TXTTAX.Text), "0.00")
End Sub

Private Sub txtws_GotFocus()
    txtWS.SelStart = 0
    txtWS.SelLength = Len(txtWS.Text)
End Sub

Private Sub txtws_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtLPack.SetFocus
        Case vbKeyEscape
            txtRT.SetFocus
    End Select
End Sub

Private Sub txtws_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtws_LostFocus()
    txtWS.Text = Format(Val(txtWS.Text), "0.00")
End Sub

Private Sub TxtHSN_GotFocus()
    txtHSN.SelStart = 0
    txtHSN.SelLength = Len(txtHSN.Text)
End Sub

Private Sub TxtHSN_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave_Click
        Case vbKeyEscape
            CmbPack.SetFocus
    End Select
End Sub

Private Sub txtHSN_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLocation_GotFocus()
    TxtLocation.SelStart = 0
    TxtLocation.SelLength = Len(TxtLocation.Text)
End Sub

Private Sub TxtLocation_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmdSave_Click
        Case vbKeyEscape
            txtLPrice.SetFocus
    End Select
End Sub

Private Sub TxtLocation_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtRT.SetFocus
        Case vbKeyEscape
            TxtBarcode.SetFocus
    End Select
End Sub

Private Sub TXTMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtMRP_LostFocus()
    TxtMRP.Text = Format(Val(TxtMRP.Text), "0.00")
End Sub

Private Sub TxtMalay_GotFocus()
    TxtMalay.SelStart = 0
    TxtMalay.SelLength = Len(TxtMalay.Text)
End Sub

Private Sub TxtMalay_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtBarcode.SetFocus
        Case vbKeyEscape
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub TxtMalay_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.Text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtMRP.SetFocus
        Case vbKeyEscape
            TxtMalay.SetFocus
    End Select
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpackdes_GotFocus()
    If Val(Txtpackdes.Text) = 0 Then
        Txtpackdes.Text = 1
    End If
    Txtpackdes.SelStart = 0
    Txtpackdes.SelLength = Len(Txtpackdes.Text)
End Sub

Private Sub Txtpackdes_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtpackdet.SetFocus
        Case vbKeyEscape
            cmbfullpack.SetFocus
    End Select
End Sub

Private Sub Txtpackdes_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtpackdet_GotFocus()
    If Val(txtpackdet.Text) = 0 Then
        txtpackdet.Text = 1
    End If
    txtpackdet.SelStart = 0
    txtpackdet.SelLength = Len(txtpackdet.Text)
End Sub

Private Sub txtpackdet_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTTAX.SetFocus
        Case vbKeyEscape
            Txtpackdes.SetFocus
    End Select
End Sub

Private Sub txtpackdet_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

