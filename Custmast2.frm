VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmcustmast 
   BackColor       =   &H00EEE3D7&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Creation"
   ClientHeight    =   9255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9450
   ControlBox      =   0   'False
   Icon            =   "Custmast2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   9450
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
      BackColor       =   &H00E0CDB8&
      Height          =   8415
      Left            =   75
      TabIndex        =   16
      Top             =   825
      Width           =   9345
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
         ItemData        =   "Custmast2.frx":16CBA
         Left            =   6045
         List            =   "Custmast2.frx":16CD3
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
         Left            =   3840
         MaxLength       =   12
         TabIndex        =   34
         Top             =   4155
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
         Left            =   1665
         MaxLength       =   25
         TabIndex        =   32
         Top             =   4155
         Width           =   990
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
         Top             =   3210
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
         Top             =   3210
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
         Height          =   390
         Left            =   8295
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   5910
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Top             =   2070
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
         Top             =   2070
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
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2820
         Width           =   4050
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
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2445
         Width           =   4050
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
         Top             =   1695
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
         Top             =   1320
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
         Top             =   1320
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
         Height          =   390
         Left            =   8295
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   5070
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   390
         Left            =   8295
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   5490
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Height          =   390
         Left            =   8295
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   6345
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin MSDataListLib.DataList Datacompany 
         Height          =   540
         Left            =   1665
         TabIndex        =   30
         Top             =   3585
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   953
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
         Left            =   45
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
         Left            =   5475
         TabIndex        =   37
         Top             =   3780
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OP. Balance"
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
         Left            =   2685
         TabIndex        =   35
         Top             =   4200
         Width           =   1350
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
         Top             =   4170
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
         Top             =   3210
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
         Top             =   2085
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GSTin No."
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
         Top             =   2085
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
         Index           =   3
         Left            =   150
         TabIndex        =   24
         Top             =   2850
         Width           =   1035
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
         Index           =   2
         Left            =   150
         TabIndex        =   23
         Top             =   2475
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
         Top             =   1725
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
         Top             =   1335
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
         Top             =   1335
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
         Height          =   675
         Left            =   1665
         TabIndex        =   4
         Top             =   630
         Width           =   6315
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   99
         BorderStyle     =   1
         Size            =   "11139;1191"
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
      MaxLength       =   15
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
Attribute VB_Name = "frmcustmast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset

Private Sub cmdcancel_Click()
    Frame.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    TXTREMARKS.Text = ""
    txtkgst.Text = ""
    TxtCST.Text = ""
    txtcompany.Text = ""
    chknewcomp.value = 0
    Txtopbal.Text = ""
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo eRRhAND
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
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
eRRhAND:
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

    On Error GoTo eRRhAND
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
    RSTITEMMAST!Remarks = Trim(TXTREMARKS.Text)
    RSTITEMMAST!KGST = Trim(txtkgst.Text)
    RSTITEMMAST!CST = Trim(TxtCST.Text)
    RSTITEMMAST!PYMT_PERIOD = Val(txtcrdtdays.Text)
    If txtcompany.Text <> "" Or Datacompany.BoundText <> "" Then
        If chknewcomp.value = 1 Then RSTITEMMAST!Area = txtcompany.Text Else RSTITEMMAST!Area = Datacompany.BoundText
    End If
    RSTITEMMAST!CONTACT_PERSON = "CS"
    RSTITEMMAST!SLSM_CODE = "SM"
    RSTITEMMAST!OPEN_DB = Round(Val(Txtopbal.Text), 3)
    RSTITEMMAST!OPEN_CR = 0
    RSTITEMMAST!YTD_DB = 0
    RSTITEMMAST!YTD_CR = 0
    RSTITEMMAST!CUST_TYPE = ""
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
    
    Frame.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    TXTREMARKS.Text = ""
    txtkgst.Text = ""
    TxtCST.Text = ""
    txtcompany.Text = ""
    Txtopbal.Text = ""
    chknewcomp.value = 0
    Txtsuplcode.Enabled = True
    CMDEXIT.Enabled = True
    CmdCancel.Enabled = True
Exit Sub
eRRhAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub Command1_Click()
    Exit Sub
    Me.Enabled = False
    frmcustTRXFILE.LBLCUSTCODE.Caption = Txtsuplcode.Text
    frmcustTRXFILE.Show
    frmcustTRXFILE.SetFocus
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRhAND
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
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    'If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    'Me.Width = 7000
    'Me.Height = 8625
    Me.Left = 2500
    Me.Top = 0
    Frame.Visible = False
    'txtunit.Visible = False
    On Error GoTo eRRhAND
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


    
    Exit Sub
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
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
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
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
        Case vbKeyReturn
            TXTREMARKS.SetFocus
        Case vbKeyEscape
            TxtCST.SetFocus
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
            TxtCST.SetFocus
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

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcompany.SetFocus
        Case vbKeyEscape
            txtdlno.SetFocus
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
        Case Asc("'"), Asc("["), Asc("]")
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
            
            On Error GoTo eRRhAND
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
                TXTREMARKS.Text = IIf(IsNull(RSTITEMMAST!Remarks), "", RSTITEMMAST!Remarks)
                txtkgst.Text = IIf(IsNull(RSTITEMMAST!KGST), "", RSTITEMMAST!KGST)
                TxtCST.Text = IIf(IsNull(RSTITEMMAST!CST), "", RSTITEMMAST!CST)
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!Area), "", RSTITEMMAST!Area)
                Txtopbal.Text = IIf(IsNull(RSTITEMMAST!OPEN_DB), 0, RSTITEMMAST!OPEN_DB)
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
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
                CmdDelete.Enabled = True
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
            Frame.Visible = True
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
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Txtsuplcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplist_Change()
    On Error GoTo eRRhAND
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
eRRhAND:
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
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub txtsupplist_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
    On Error GoTo eRRhAND
    
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
eRRhAND:
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
            TXTREMARKS.SetFocus
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
            On Error GoTo eRRhAND
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
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub Datacompany_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo eRRhAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEM_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        txtcompany.Text = RSTITEMMAST!MANUFACTURER
'    End If
    txtcompany.Text = Datacompany.BoundText
    Datacompany.Text = txtcompany.Text
    Exit Sub
eRRhAND:
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
            Txtopbal.SetFocus
        Case vbKeyEscape
            Datacompany.SetFocus
    End Select
End Sub

Private Sub txtcrdtdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
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
            txtcrdtdays.SetFocus
    End Select
End Sub

Private Sub Txtopbal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
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
            CmdSave.SetFocus
        Case vbKeyEscape
            Txtopbal.SetFocus
    End Select
End Sub
