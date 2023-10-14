VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSTKSUMRy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK SUMMARY"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19155
   ClipControls    =   0   'False
   Icon            =   "FRMSTKSUM2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   19155
   Begin VB.CheckBox chkshowsup 
      Caption         =   "Show last supplier"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13515
      TabIndex        =   45
      Top             =   8145
      Width           =   1950
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print Price List"
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
      Left            =   17430
      TabIndex        =   44
      Top             =   6960
      Width           =   1230
   End
   Begin VB.CheckBox ChkUnBill 
      Caption         =   "Show Un Bill Items"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   15495
      TabIndex        =   43
      Top             =   8160
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   1095
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Left            =   870
      ScaleHeight     =   240
      ScaleWidth      =   2070
      TabIndex        =   39
      Top             =   -15
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Left            =   13800
      ScaleHeight     =   240
      ScaleWidth      =   2070
      TabIndex        =   38
      Top             =   8910
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
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
      Height          =   210
      Left            =   -345
      ScaleHeight     =   210
      ScaleWidth      =   390
      TabIndex        =   37
      Top             =   255
      Visible         =   0   'False
      Width           =   390
   End
   Begin VB.CommandButton CMDPRINTlABELS 
      Caption         =   "PRINT &LABELS"
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
      Left            =   13485
      TabIndex        =   36
      Top             =   7410
      Width           =   1275
   End
   Begin VB.CheckBox ChkDetails 
      Caption         =   "Detailed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   13500
      TabIndex        =   14
      Top             =   7860
      Width           =   1365
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&PRINT"
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
      Left            =   14820
      TabIndex        =   13
      Top             =   6960
      Width           =   1305
   End
   Begin VB.CommandButton CmdDisplay 
      Caption         =   "&Display"
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
      Left            =   13470
      TabIndex        =   12
      Top             =   6960
      Width           =   1290
   End
   Begin VB.Frame Frame 
      Height          =   2190
      Left            =   3750
      TabIndex        =   8
      Top             =   2970
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2640
         TabIndex        =   2
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   1335
         TabIndex        =   1
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Frame Frame1 
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
         TabIndex        =   9
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
            TabIndex        =   42
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
            TabIndex        =   41
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
            TabIndex        =   0
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
            TabIndex        =   10
            Top             =   765
            Width           =   1260
         End
      End
   End
   Begin VB.TextBox TXTsample 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   6195
      TabIndex        =   5
      Top             =   1545
      Visible         =   0   'False
      Width           =   1350
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
      Left            =   16155
      TabIndex        =   4
      Top             =   6960
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   6930
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   19080
      _ExtentX        =   33655
      _ExtentY        =   12224
      _Version        =   393216
      Rows            =   1
      Cols            =   24
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frmeall 
      BackColor       =   &H00FFC0C0&
      Height          =   1620
      Left            =   0
      TabIndex        =   11
      Top             =   6855
      Width           =   13470
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by......"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   6840
         TabIndex        =   28
         Top             =   120
         Width           =   4530
         Begin VB.OptionButton OptCode 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Item Code"
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
            Height          =   315
            Left            =   75
            TabIndex        =   30
            Top             =   630
            Width           =   2070
         End
         Begin VB.OptionButton OptLow 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Low Price"
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
            Height          =   315
            Left            =   2175
            TabIndex        =   35
            Top             =   1155
            Width           =   2160
         End
         Begin VB.OptionButton OptHighest 
            BackColor       =   &H00FFC0C0&
            Caption         =   "High Price"
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
            Height          =   315
            Left            =   2175
            TabIndex        =   34
            Top             =   825
            Width           =   2160
         End
         Begin VB.OptionButton OptDead 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Dead moving items"
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
            Height          =   315
            Left            =   2175
            TabIndex        =   32
            Top             =   195
            Width           =   2160
         End
         Begin VB.OptionButton Optfast 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Fast moving Items"
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
            Height          =   315
            Left            =   2175
            TabIndex        =   33
            Top             =   525
            Width           =   2085
         End
         Begin VB.OptionButton OptCategory 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Category"
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
            Height          =   315
            Left            =   75
            TabIndex        =   31
            Top             =   1005
            Width           =   2325
         End
         Begin VB.OptionButton OptSortName 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Name"
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
            Height          =   315
            Left            =   75
            TabIndex        =   29
            Top             =   285
            Value           =   -1  'True
            Width           =   2070
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   1410
         Left            =   11385
         TabIndex        =   25
         Top             =   180
         Width           =   2040
         Begin VB.OptionButton OptStock 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Stock Items Only"
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
            Height          =   315
            Left            =   90
            TabIndex        =   27
            Top             =   750
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Display All"
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
            Height          =   315
            Left            =   75
            TabIndex        =   26
            Top             =   300
            Width           =   1335
         End
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
         Left            =   3525
         TabIndex        =   21
         Top             =   435
         Width           =   3300
      End
      Begin VB.CheckBox CHKCATEGORY2 
         BackColor       =   &H00FFC0C0&
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
         Left            =   3510
         TabIndex        =   20
         Top             =   120
         Width           =   1335
      End
      Begin VB.CheckBox CHKCATEGORY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Manufacturer"
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
         Left            =   180
         TabIndex        =   19
         Top             =   120
         Width           =   1680
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   180
         TabIndex        =   15
         Top             =   435
         Width           =   3285
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   780
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   3285
         _ExtentX        =   5794
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
      Begin MSDataListLib.DataList DataList1 
         Height          =   780
         Left            =   3525
         TabIndex        =   22
         Top             =   780
         Width           =   3300
         _ExtentX        =   5821
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
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Index           =   0
      Left            =   17325
      TabIndex        =   47
      Top             =   7395
      Width           =   1215
   End
   Begin VB.Label lblnetvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   405
      Left            =   16905
      TabIndex        =   46
      Top             =   7650
      Width           =   2145
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   24
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   23
      Top             =   480
      Width           =   495
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   14865
      TabIndex        =   7
      Top             =   7650
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Net Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   14910
      TabIndex        =   6
      Top             =   7395
      Width           =   1830
   End
End
Attribute VB_Name = "FRMSTKSUMRy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Dim PHY_REC As New ADODB.Recordset
Dim PHY_FLAG As Boolean
Dim printerfound As Boolean

Private Sub CHKCATEGORY2_Click()
    If CHKCATEGORY2.Value = 0 Then
        TXTDEALER2.Text = ""
    Else
        TXTDEALER2.SetFocus
    End If
End Sub

Private Sub CHKCATEGORY_Click()
    If chkcategory.Value = 0 Then
        TXTDEALER.Text = ""
    Else
        TXTDEALER.SetFocus
    End If
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

Private Sub CmDDisplay_Click()
    If chkcategory.Value = 1 And DataList2.BoundText = "" Then
        MsgBox "Select Manufacturer from the List", vbOKOnly, "Stock Register"
        DataList2.SetFocus
        Exit Sub
    End If
    
    If CHKCATEGORY2.Value = 1 And DataList1.BoundText = "" Then
        MsgBox "Select Category from the List", vbOKOnly, "Stock Register"
        DataList1.SetFocus
        Exit Sub
    End If
    
    If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
        Call Fillgrid
    Else
        Call Fillgrid
    End If
    GRDSTOCK.SetFocus
End Sub

Private Sub CmdExit_Click()
    Unload Me
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
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = "0.00"
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18) = ""
        Else
            If OptAmt.Value = True Then
                rststock!COM_FLAG = "A"
                rststock!COM_PER = 0
                rststock!COM_AMT = Val(TxtComper.Text)
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18) = "Rs"
            Else
                rststock!COM_FLAG = "P"
                rststock!COM_PER = Val(TxtComper.Text)
                rststock!COM_AMT = 0
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 17) = Format(Val(TxtComper.Text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18) = "%"
            End If
        End If
        rststock.Update
    End If
    rststock.Close
    Set rststock = Nothing
    GRDSTOCK.Enabled = True
    FRAME.Visible = False
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim i As Long
    
    On Error GoTo eRRhAND
    
    If ChkDetails.Value = 0 Then
        ReportNameVar = Rptpath & "RPTSTOCKSMRY"
    Else
        If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then
            ReportNameVar = Rptpath & "RPTSTOCKS"
        Else
            ReportNameVar = Rptpath & "RPTSTOCK_DET"
        End If
    End If
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {ITEMMAST.CLOSE_QTY}<> 0 )"
    If chkunbill.Value = 1 Then
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        End If
    Else
        'ISNULL(UN_BILL) OR UN_BILL <> 'Y'
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
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
eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    
    On Error GoTo eRRhAND
    ReportNameVar = Rptpath & "RPTPRICELIST"
    
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {ITEMMAST.CLOSE_QTY}<> 0 )"
    If chkunbill.Value = 1 Then
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        End If
    Else
        'ISNULL(UN_BILL) OR UN_BILL <> 'Y'
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.MANUFACTURER} = '" & DataList2.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' )"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} = '" & DataList1.BoundText & "' AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
              Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES')"
            Else
                Report.RecordSelectionFormula = "((ISNULL({ITEMMAST.DEAD_STOCK}) OR {ITEMMAST.DEAD_STOCK} <> 'Y') AND (ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND {ITEMMAST.CATEGORY} <> 'SERVICE CHARGE' AND {ITEMMAST.CATEGORY} <> 'SERVICES' AND {ITEMMAST.CLOSE_QTY}<> 0 )"
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
        ''If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.Text = "'A R STEELS' & chr(13) & 'Alappuzha'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'Price List'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "STOCK REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    ACT_FLAG = True
    PHY_FLAG = True
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "Total Qty"
    GRDSTOCK.TextMatrix(0, 4) = "PACK"
    GRDSTOCK.TextMatrix(0, 5) = "MRP"
    GRDSTOCK.TextMatrix(0, 6) = "RT"
    GRDSTOCK.TextMatrix(0, 7) = "WS"
    GRDSTOCK.TextMatrix(0, 8) = "TAX"
    GRDSTOCK.TextMatrix(0, 9) = "Rcvd Qty"
    GRDSTOCK.TextMatrix(0, 10) = "Issu. Qty"
    GRDSTOCK.TextMatrix(0, 11) = "Cost"
    GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
    GRDSTOCK.TextMatrix(0, 13) = "Net Value"
    GRDSTOCK.TextMatrix(0, 14) = "" '"TRX TYPE"
    GRDSTOCK.TextMatrix(0, 15) = "" '"VCH NO"
    GRDSTOCK.TextMatrix(0, 16) = "" '"LINE NO"
    GRDSTOCK.TextMatrix(0, 17) = "COMISSION"
    GRDSTOCK.TextMatrix(0, 18) = "TYPE"
    GRDSTOCK.TextMatrix(0, 19) = "Category"
    GRDSTOCK.TextMatrix(0, 20) = "Qty"
    GRDSTOCK.TextMatrix(0, 21) = "Loose Pack"
    GRDSTOCK.TextMatrix(0, 22) = "L.Price"
    GRDSTOCK.TextMatrix(0, 23) = "Last Bill & Suppiler"
    
    GRDSTOCK.ColWidth(0) = 700
    'GRDSTOCK.ColWidth(1) = 1500
    GRDSTOCK.ColWidth(2) = 3000
    GRDSTOCK.ColWidth(3) = 950
    GRDSTOCK.ColWidth(4) = 950
    GRDSTOCK.ColWidth(5) = 0
    If frmLogin.rs!Level <> "0" Then
        GRDSTOCK.ColWidth(16) = 0
    Else
        GRDSTOCK.ColWidth(16) = 1100
    End If
    GRDSTOCK.ColWidth(7) = 1100
    GRDSTOCK.ColWidth(8) = 800
    GRDSTOCK.ColWidth(9) = 0
    GRDSTOCK.ColWidth(10) = 0
    GRDSTOCK.ColWidth(11) = 1100
    GRDSTOCK.ColWidth(12) = 1100
    GRDSTOCK.ColWidth(13) = 1100
    GRDSTOCK.ColWidth(14) = 0
    GRDSTOCK.ColWidth(15) = 0
    GRDSTOCK.ColWidth(16) = 0
    GRDSTOCK.ColWidth(18) = 600
    GRDSTOCK.ColWidth(19) = 1200
    GRDSTOCK.ColWidth(20) = 0
    GRDSTOCK.ColWidth(21) = 0
    GRDSTOCK.ColWidth(22) = 1000
    GRDSTOCK.ColWidth(23) = 2800
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 1
'    GRDSTOCK.ColAlignment(5) = 1
'    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(18) = 1
    GRDSTOCK.ColAlignment(19) = 1
    
    Picture2.ScaleMode = 3
    Picture2.Height = Picture2.Height * (1.4 * 40 / Picture2.ScaleHeight)
    Picture2.FontSize = 8
    Picture1.FontSize = 5
    Picture3.FontSize = 5
    Picture4.FontSize = 5
    
    Me.Left = 0
    Me.Top = 0
    
    If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
        GRDSTOCK.ColWidth(9) = 0 '1100
        GRDSTOCK.ColWidth(10) = 0 ' 1100
        Call Fillgrid
    Else
        GRDSTOCK.ColWidth(9) = 0
        GRDSTOCK.ColWidth(10) = 0
        GRDSTOCK.ColWidth(11) = 0
        GRDSTOCK.ColWidth(12) = 0
        GRDSTOCK.ColWidth(13) = 0
        lblpvalue.Visible = False
        lblnetvalue.Visible = False
        Call Fillgrid
    End If
    'Me.Height = 10000
    'Me.Width = 14595
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus

'    MDIMAIN.PCTMENU.Enabled = True
'    'MDIMAIN.PCTMENU.Height = 555
'    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDSTOCK.Col
                    Case 2, 5, 6, 7, 8, 11, 19, 21, 22
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop '+ 50
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 50
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 17
                        FRAME.Visible = True
                        FRAME.Top = GRDSTOCK.CellTop - 800
                        FRAME.Left = GRDSTOCK.CellLeft - 1500
                        'Frame.Width = GRDSTOCK.CellWidth - 25
                        TxtComper.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 18) = "Rs" Then
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
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
    FRAME.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub lblpvalue_DblClick()
    If chkunbill.Visible = True Then
        chkunbill.Value = 0
        chkunbill.Visible = False
    Else
        chkunbill.Value = 0
        chkunbill.Visible = True
    End If
End Sub

Private Sub OptAmt_Click()
    TxtComper.SetFocus
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

Private Sub OptPercent_Click()
    TxtComper.SetFocus
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

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                
'                Case 1  ' Item Code
'                    If Trim(TXTsample.Text) = "" Then Exit Sub
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!ITEM_CODE = Trim(TXTsample.Text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    Do Until rststock.EOF
'                        rststock!ITEM_CODE = Trim(TXTsample.Text)
'                        rststock.Update
'                        rststock.MoveNext
'                    Loop
'                    rststock.Close
'                    Set rststock = Nothing
'
'                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
'                    GRDSTOCK.Enabled = True
'                    TXTsample.Visible = False
'                    GRDSTOCK.SetFocus
                    
                Case 2  ' Item Name
                    If Trim(TXTsample.Text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_NAME = Trim(TXTsample.Text)
                        rststock.Update
                    End If
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
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 3  ' Bal QTY
                    M_STOCK = 0
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!CLOSE_QTY = Val(TXTsample.Text)
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / 100)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.000")
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 6  'RT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 7  'WS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 11  'COST
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!item_COST = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(TXTsample.Text) * rststock!CLOSE_QTY, "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / 100)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.000")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format((Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) + (Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) / 100)), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 8  'TAX
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!SALES_TAX = Val(TXTsample.Text)
                        rststock!check_flag = "V"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 22  'CRTN
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_CRTN = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 21  'CRTN PCK
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CRTN_PACK = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                                
                Case 19  'CATEGORY
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!Category = Trim(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!Category = Trim(TXTsample.Text)
                        rststock.Update
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 5  'MRP
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.Text)
                        'rststock!P_RETAIL = Val(TXTsample.Text)
                        'rststock!P_WS = Val(TXTsample.Text)
                        'rststock!P_VAN = Val(TXTsample.Text)
                        
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(rststock!MRP, "0.000")
                        
                        ''Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * 15 / 100, ".000")
                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4)), 2)
                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) - Val(TXTsample.Text) * 15 / 100, 2)
                        ''grdsTOCK.TextMatrix(grdsTOCK.Row, 8) = Format(rststock!P_RETAIL, "0.00")
                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format(rststock!P_RETAIL * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
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

Private Function STOCKADJUST()
'    Dim rststock As ADODB.Recordset
'    Dim RSTITEMMAST As ADODB.Recordset
'
'
'    On Error GoTo eRRHAND
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT CLOSE_QTY from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly
'    Do Until rststock.EOF
'        M_STOCK = M_STOCK + rststock!CLOSE_QTY
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
'
'
'    Exit Function
'
'eRRHAND:
'    MsgBox Err.Description
End Function

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 5, 6, 7, 8, 9, 11, 21, 22
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 19
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
'        Case 5
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
'        Case 7
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
    End Select
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Function TOTALVALUE()
    Dim i As Long
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
    For i = 1 To GRDSTOCK.rows - 1
        lblpvalue.Caption = Val(lblpvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 3)))
        lblnetvalue.Caption = Val(lblnetvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 13))
    Next i
    lblpvalue.Caption = Format(lblpvalue.Caption, "0.00")
    lblnetvalue.Caption = Format(lblnetvalue.Caption, "0.00")
End Function

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

Private Sub TXTDEALER_Change()
    
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!MANUFACTURER
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "MANUFACTURER"
        DataList2.BoundColumn = "MANUFACTURER"
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
    chkcategory.Value = 1
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
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

Private Sub DataList2_Click()
        
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text

End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER.Text) = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
    chkcategory.Value = 1
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Fillgrid()
    Dim rststock As ADODB.Recordset
 
    Dim i As Long
    Dim P_Value As Double
    Dim S_Value As Double
    
    
    On Error GoTo eRRhAND
    i = 0
        
    Screen.MousePointer = vbHourglass
    
    S_Value = 0
    P_Value = 0
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.rows = 1
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM RTRXFILE WHERE  RTRXFILE.CLOSE_QTY > 0 ORDER BY RTRXFILE.ITEM_NAME", DB, adOpenStatic,adLockReadOnly
    
    If chkunbill.Value = 0 Then
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N') AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        End If
        '---------------
    Else
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList2.BoundText & "' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE MANUFACTURER = '" & DataList2.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE CATEGORY = '" & DataList1.BoundText & "' AND ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            Else
                If OptSortName.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                ElseIf OptCode.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY CONVERT(ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                ElseIf optCategory.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                ElseIf OptDead.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                ElseIf Optfast.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                ElseIf OptLow.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                ElseIf OptHighest.Value = True Then
                    rststock.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND UCASE(CATEGORY) <> 'SERVICE CHARGE'  AND CATEGORY <> 'SERVICES' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                End If
            End If
        End If
    End If
    Dim RSTSUPPLIER As ADODB.Recordset
    Do Until rststock.EOF
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = Round(rststock!CLOSE_QTY, 3)
        If IsNull(rststock!LOOSE_PACK) Then
            GRDSTOCK.TextMatrix(i, 4) = 1
        Else
            'GRDSTOCK.TextMatrix(i, 4) = rststock!LOOSE_PACK & IIf(IsNull(rststock!PACK_TYPE), "", " " & rststock!PACK_TYPE)
            GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!PACK_TYPE), "", " " & rststock!PACK_TYPE)
        End If
        'GRDSTOCK.TextMatrix(i, 3) = Val(GRDSTOCK.TextMatrix(i, 3)) / IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!P_RETAIL), "", Format(rststock!P_RETAIL, "0.000"))
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_WS), "", Format(rststock!P_WS, "0.000"))
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!SALES_TAX), "", Format(rststock!SALES_TAX, "0.000"))
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!RCPT_QTY), "", rststock!RCPT_QTY)
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!ISSUE_QTY), "", rststock!ISSUE_QTY)
        GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!item_COST), "", Format(rststock!item_COST, "0.000"))
        GRDSTOCK.TextMatrix(i, 12) = Format((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 8)) / 100)), "0.000")
        GRDSTOCK.TextMatrix(i, 13) = Format((Val(GRDSTOCK.TextMatrix(i, 11)) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 8)) / 100)) * Val(GRDSTOCK.TextMatrix(i, 3)), "0.000")
        lblpvalue.Caption = Val(lblpvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 11)) * Val(GRDSTOCK.TextMatrix(i, 3)))
        lblnetvalue.Caption = Val(lblnetvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 13))
        
        Select Case rststock!COM_FLAG
            Case "P"
                GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 18) = "%"
            Case "A"
                GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                GRDSTOCK.TextMatrix(i, 18) = "Rs"
        End Select
        GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!Category), "", rststock!Category)
        Select Case rststock!LOOSE_PACK
            Case Is > 1
                If Int(rststock!CLOSE_QTY) <= 1 Then
                    GRDSTOCK.TextMatrix(i, 20) = Int(rststock!CLOSE_QTY) & " Packet" & IIf(Val(Round((rststock!CLOSE_QTY - Int(rststock!CLOSE_QTY)) * rststock!LOOSE_PACK, 0)) = 0, "", " & " & Round((rststock!CLOSE_QTY - Int(rststock!CLOSE_QTY)) * rststock!LOOSE_PACK, 0) & " " & rststock!PACK_TYPE)
                Else
                    GRDSTOCK.TextMatrix(i, 20) = Int(rststock!CLOSE_QTY) & " Packets" & IIf(Val(Round((rststock!CLOSE_QTY - Int(rststock!CLOSE_QTY)) * rststock!LOOSE_PACK, 0)) = 0, "", " & " & Round((rststock!CLOSE_QTY - Int(rststock!CLOSE_QTY)) * rststock!LOOSE_PACK, 0) & " " & rststock!PACK_TYPE)
                    'GRDSTOCK.TextMatrix(i, 18) = Int(rststock!CLOSE_QTY) & " Packets & " & Round((rststock!CLOSE_QTY - Int(rststock!CLOSE_QTY)) * rststock!LOOSE_PACK, 0) & " " & rststock!PACK_TYPE
                End If
                
            Case Else
                GRDSTOCK.TextMatrix(i, 20) = Int(rststock!CLOSE_QTY)
        End Select
        GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!CRTN_PACK), "", rststock!CRTN_PACK)
        GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!P_CRTN), "", Format(rststock!P_CRTN, "0.00"))
        'grdsTOCK.TextMatrix(i, 23) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
        
        If chkshowsup.Value = 1 Then
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT VCH_NO, VCH_DESC, TRX_TYPE FROM RTRXFILE WHERE (TRX_TYPE = 'PI' or TRX_TYPE = 'PW') AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  TRX_TYPE, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                Select Case RSTSUPPLIER!TRX_TYPE
                    Case "PI"
                        GRDSTOCK.TextMatrix(i, 23) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                    Case Else
                        GRDSTOCK.TextMatrix(i, 23) = "W- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                End Select
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
        End If
        rststock.MoveNext
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
    Loop
    rststock.Close
    Set rststock = Nothing

    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub TXTDEALER2_Change()
    
    On Error GoTo eRRhAND
    If flagchange2.Caption <> "1" Then
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
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
    CHKCATEGORY2.Value = 1
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text

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
            CMDDISPLAY.SetFocus
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
    CHKCATEGORY2.Value = 1
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

