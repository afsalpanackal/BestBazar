VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLenders 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lender Master"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6795
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   6795
   Begin VB.TextBox txtsupplist 
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
      Left            =   1680
      MaxLength       =   34
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H0080C0FF&
      Height          =   5190
      Left            =   120
      TabIndex        =   16
      Top             =   1245
      Width           =   6660
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
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   31
         Top             =   1515
         Width           =   2895
      End
      Begin VB.CheckBox chknewcomp 
         BackColor       =   &H00800080&
         Caption         =   "&New Area"
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
         Height          =   255
         Left            =   4530
         TabIndex        =   30
         Top             =   1875
         Width           =   1695
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
         Height          =   345
         Left            =   1590
         MaxLength       =   12
         TabIndex        =   28
         Top             =   3945
         Width           =   2235
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
         Left            =   120
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   4605
         UseMaskColor    =   -1  'True
         Width           =   1260
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
         Left            =   1590
         MaxLength       =   25
         TabIndex        =   11
         Top             =   8580
         Visible         =   0   'False
         Width           =   2235
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
         Left            =   1590
         MaxLength       =   25
         TabIndex        =   10
         Top             =   7320
         Visible         =   0   'False
         Width           =   2235
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
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   9
         Top             =   7815
         Visible         =   0   'False
         Width           =   4230
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
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   8
         Top             =   7440
         Visible         =   0   'False
         Width           =   4230
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
         Left            =   1590
         MaxLength       =   50
         TabIndex        =   7
         Top             =   3540
         Width           =   4215
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   6
         Top             =   3165
         Width           =   2235
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2790
         Width           =   2235
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
         Left            =   1590
         MaxLength       =   34
         TabIndex        =   3
         Top             =   225
         Width           =   4980
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
         Left            =   3510
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   4605
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
         Left            =   4545
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   4605
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
         Left            =   5580
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   4605
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin MSDataListLib.DataList Datacompany 
         Height          =   780
         Left            =   1590
         TabIndex        =   32
         Top             =   1890
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         Left            =   75
         TabIndex        =   33
         Top             =   1710
         Width           =   1035
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
         Index           =   6
         Left            =   150
         TabIndex        =   29
         Top             =   3975
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CST No."
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
         Left            =   150
         TabIndex        =   26
         Top             =   8610
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN No."
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
         Top             =   8235
         Visible         =   0   'False
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
         Left            =   150
         TabIndex        =   24
         Top             =   7845
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KGST NO. 1"
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
         Top             =   7470
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
         Top             =   3570
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Fax No."
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
         Left            =   150
         TabIndex        =   21
         Top             =   3195
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
         Top             =   2820
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
         Height          =   855
         Left            =   1590
         TabIndex        =   4
         Top             =   570
         Width           =   4980
         VariousPropertyBits=   746604571
         ForeColor       =   255
         MaxLength       =   99
         BorderStyle     =   1
         Size            =   "8784;1508"
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
         Caption         =   "Lender Name"
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
         Left            =   150
         TabIndex        =   17
         Top             =   255
         Width           =   1365
      End
   End
   Begin VB.TextBox Txtsuplcode 
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
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1680
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2858
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
      Caption         =   "Lender Code"
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
Attribute VB_Name = "frmLenders"
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
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtremarks.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    Txtopbal.Text = ""
    txtcompany.Text = ""
    chknewcomp.value = 0
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From RTRXFILE WHERE M_USER_ID = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRANSMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "SALES") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    db.Execute ("delete  FROM PRODLINK WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME OF SUPPLIER", vbOKOnly, "LENDER MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If
    
    If chknewcomp.value = 0 And Datacompany.BoundText = "" And txtcompany.Text <> "" Then
        MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "PRODUCT MASTER"
        txtcompany.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
    Txtsuplcode.Tag = "211" & Txtsuplcode.Text
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
    
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
        RSTITEMMAST!Address = Trim(txtaddress.Text)
        RSTITEMMAST!TELNO = Trim(txttelno.Text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
        RSTITEMMAST!DL_NO = Trim(txtdlno.Text)
        RSTITEMMAST!Remarks = Trim(txtremarks.Text)
        RSTITEMMAST!KGST = Trim(txtkgst.Text)
        RSTITEMMAST!CST = Trim(txtcst.Text)
        RSTITEMMAST!OPEN_DB = Round(Val(Txtopbal.Text), 3)
        If txtcompany.Text <> "" Or Datacompany.BoundText <> "" Then
            If chknewcomp.value = 1 Then RSTITEMMAST!Area = txtcompany.Text Else RSTITEMMAST!Area = Datacompany.BoundText
        End If
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!ACT_CODE = Txtsuplcode.Tag
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
        RSTITEMMAST!Address = Trim(txtaddress.Text)
        RSTITEMMAST!TELNO = Trim(txttelno.Text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
        RSTITEMMAST!DL_NO = Trim(txtdlno.Text)
        RSTITEMMAST!Remarks = Trim(txtremarks.Text)
        RSTITEMMAST!KGST = Trim(txtkgst.Text)
        RSTITEMMAST!CST = Trim(txtcst.Text)
        
        RSTITEMMAST!Area = ""
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
        If txtcompany.Text <> "" Or Datacompany.BoundText <> "" Then
            If chknewcomp.value = 1 Then RSTITEMMAST!Area = txtcompany.Text Else RSTITEMMAST!Area = Datacompany.BoundText
        End If

        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtremarks.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    Txtopbal.Text = ""
    Txtsuplcode.Enabled = True
    txtcompany.Text = ""
    chknewcomp.value = 0
Exit Sub
ErrHand:
    MsgBox (Err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Txtopbal.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ACT_CODE FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.Text = Mid(RSTITEMMAST!ACT_CODE, Len(RSTITEMMAST!ACT_CODE) - 2, 3)
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
    If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    'Me.Width = 6930
    'Me.Height = 8265
    Me.Left = 3500
    Me.Top = 100
    FRAME.Visible = False
    'txtunit.Visible = False
    On Error GoTo ErrHand
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(ACT_CODE) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3)", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "001", Mid(TRXMAST.Fields(0), 4) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
ErrHand:
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
            txtcompany.SetFocus
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
            Txtopbal.SetFocus
    End Select
End Sub

Private Sub txtdlno_GotFocus()
    txtdlno.SelStart = 0
    txtdlno.SelLength = Len(txtdlno.Text)
End Sub

Private Sub txtdlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
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
            Txtopbal.SetFocus
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
    End Select
End Sub

Private Sub Txtopbal_GotFocus()
    Txtopbal.SelStart = 0
    Txtopbal.SelLength = Len(Txtopbal.Text)
End Sub

Private Sub Txtopbal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
        Case vbKeyEscape
            txtemail.SetFocus
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

Private Sub txtremarks_GotFocus()
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtkgst.SetFocus
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
                MsgBox "ENTER NAME OF LENDER", vbOKOnly, "LENDER MASTER"
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
            If Val(Txtsuplcode.Text) = 0 Then Exit Sub
            On Error GoTo ErrHand
            Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
            Txtsuplcode.Tag = "211" & Txtsuplcode.Text
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.Text = RSTITEMMAST!ACT_NAME
                txtaddress.Text = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                txttelno.Text = IIf(IsNull(RSTITEMMAST!TELNO), "", RSTITEMMAST!TELNO)
                txtfaxno.Text = IIf(IsNull(RSTITEMMAST!FAXNO), "", RSTITEMMAST!FAXNO)
                txtemail.Text = IIf(IsNull(RSTITEMMAST!EMAIL_ADD), "", RSTITEMMAST!EMAIL_ADD)
                txtdlno.Text = IIf(IsNull(RSTITEMMAST!DL_NO), "", RSTITEMMAST!DL_NO)
                txtremarks.Text = IIf(IsNull(RSTITEMMAST!Remarks), "", RSTITEMMAST!Remarks)
                txtkgst.Text = IIf(IsNull(RSTITEMMAST!KGST), "", RSTITEMMAST!KGST)
                txtcst.Text = IIf(IsNull(RSTITEMMAST!CST), "", RSTITEMMAST!CST)
                Txtopbal.Text = IIf(IsNull(RSTITEMMAST!OPEN_DB), 0, RSTITEMMAST!OPEN_DB)
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!Area), "", RSTITEMMAST!Area)
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
                CmdDelete.Enabled = True
            End If
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub txtsupplist_Change()
    On Error GoTo ErrHand
    If REPFLAG = True Then
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
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
            txtcompany.SetFocus
    End Select
End Sub

Private Sub txtcompany_Change()
    On Error GoTo ErrHand
    
    Set Me.Datacompany.RowSource = Nothing
    If COMPANYFLAG = True Then
        RSTCOMPANY.Open "Select DISTINCT AREA From ACTMAST WHERE AREA Like '" & txtcompany.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    Else
        RSTCOMPANY.Close
        RSTCOMPANY.Open "Select DISTINCT AREA From ACTMAST WHERE AREA Like '" & txtcompany.Text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
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
            txtaddress.SetFocus
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
            RSTITEMMAST.Open "SELECT AREA FROM ACTMAST WHERE ACT_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtcompany.Text = RSTITEMMAST!Area
            Else
'                If txtcompany.Text = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
'                If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "PRODUCT MASTER"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
                If chknewcomp.value = 0 And Datacompany.BoundText = "" And txtcompany.Text <> "" Then
                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "PRODUCT MASTER"
                    txtcompany.SetFocus
                    Exit Sub
                End If
            End If
            txttelno.SetFocus
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

