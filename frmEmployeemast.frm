VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmEmployeeMast 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Master"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6915
   ControlBox      =   0   'False
   Icon            =   "frmEmployeemast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   6915
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
      Height          =   5310
      Left            =   120
      TabIndex        =   13
      Top             =   1245
      Width           =   6660
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
         TabIndex        =   21
         Top             =   4815
         UseMaskColor    =   -1  'True
         Width           =   1260
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
         Left            =   1560
         MaxLength       =   49
         TabIndex        =   8
         Top             =   3105
         Width           =   2235
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
         Left            =   1560
         MaxLength       =   25
         TabIndex        =   7
         Top             =   2625
         Width           =   3750
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   6
         Top             =   2130
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
         Left            =   1560
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1575
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
         TabIndex        =   9
         Top             =   4815
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
         TabIndex        =   10
         Top             =   4815
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
         TabIndex        =   11
         Top             =   4815
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   90
         TabIndex        =   20
         Top             =   3135
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
         Left            =   105
         TabIndex        =   19
         Top             =   2655
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mob No."
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
         Left            =   120
         TabIndex        =   18
         Top             =   2160
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
         Left            =   75
         TabIndex        =   17
         Top             =   1605
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
         Left            =   60
         TabIndex        =   16
         Top             =   645
         Width           =   1470
      End
      Begin MSForms.TextBox txtaddress 
         Height          =   855
         Left            =   1575
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
         Caption         =   "Employee Name"
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
         Left            =   60
         TabIndex        =   14
         Top             =   255
         Width           =   1545
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
      TabIndex        =   15
      Top             =   45
      Width           =   6300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Employee Code"
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
      TabIndex        =   12
      Top             =   465
      Width           =   1560
   End
End
Attribute VB_Name = "frmEmployeeMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
'Dim rstTMP As New ADODB.Recordset
'Dim TMPFLAG As Boolean 'TMP

Private Sub cmdcancel_Click()
    Frame.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    TXTREMARKS.Text = ""
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo eRRHAND
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From RTRXFILE WHERE M_USER_ID = '" & Txtsuplcode.Tag & "'", db, adOpenForwardOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRANSMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenForwardOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenForwardOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "SALES") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME OF SUPPLIER", vbOKOnly, "CUSTOMER MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If

    On Error GoTo eRRHAND
    Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
    Txtsuplcode.Tag = "321" & Txtsuplcode.Text
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
    
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
        RSTITEMMAST!Address = Trim(txtaddress.Text)
        RSTITEMMAST!TELNO = Trim(txttelno.Text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
        RSTITEMMAST!DL_NO = ""
        RSTITEMMAST!Remarks = Trim(TXTREMARKS.Text)
        RSTITEMMAST!KGST = ""
        RSTITEMMAST!CST = ""
        RSTITEMMAST!OPEN_DB = 0
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!act_code = Txtsuplcode.Tag
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
        RSTITEMMAST!Address = Trim(txtaddress.Text)
        RSTITEMMAST!TELNO = Trim(txttelno.Text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
        RSTITEMMAST!DL_NO = ""
        RSTITEMMAST!Remarks = Trim(TXTREMARKS.Text)
        RSTITEMMAST!KGST = ""
        RSTITEMMAST!CST = ""
        
        RSTITEMMAST!Area = "SM"
        RSTITEMMAST!CONTACT_PERSON = "CS"
        RSTITEMMAST!SLSM_CODE = "SM"
        RSTITEMMAST!OPEN_DB = 0
        RSTITEMMAST!OPEN_CR = 0
        RSTITEMMAST!YTD_DB = 0
        RSTITEMMAST!YTD_CR = 0
        RSTITEMMAST!CUST_TYPE = ""
        RSTITEMMAST!CREATE_DATE = Date
        RSTITEMMAST!C_USER_ID = "SM"
        RSTITEMMAST!MODIFY_DATE = Date
        RSTITEMMAST!M_USER_ID = "SM"


        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    Frame.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    TXTREMARKS.Text = ""
    Txtsuplcode.Enabled = True
Exit Sub
eRRHAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRHAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ACT_CODE FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.Text = Mid(RSTITEMMAST!act_code, Len(RSTITEMMAST!act_code) - 2, 3)
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    'TMPFLAG = True
    Width = 8700
    Height = 7500
    Left = 3500
    Top = 1900
    Frame.Visible = False
    'txtunit.Visible = False
    On Error GoTo eRRHAND
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(ACT_CODE) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='321')And (LENGTH(ACT_CODE)>3)", db, adOpenForwardOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "001", Mid(TRXMAST.Fields(0), 4) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub txtaddress_GotFocus()
    txtaddress.SelStart = 0
    txtaddress.SelLength = Len(txtaddress.Text)
End Sub

Private Sub txtaddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txttelno.SetFocus
    End Select
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtemail_GotFocus()
    txtemail.SelStart = 0
    txtemail.SelLength = Len(txtemail.Text)
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub txtemail_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("@")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
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
    End Select
End Sub

Private Sub txtfaxno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
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
            CmdSave.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
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
                MsgBox "ENTER NAME OF SUPPLIER", vbOKOnly, "PRODUCT MASTER"
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
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
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
            On Error GoTo eRRHAND
            Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
            Txtsuplcode.Tag = "321" & Txtsuplcode.Text
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenForwardOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.Text = IIf(IsNull(RSTITEMMAST!ACT_NAME), "", RSTITEMMAST!ACT_NAME)
                txtaddress.Text = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                txttelno.Text = IIf(IsNull(RSTITEMMAST!TELNO), "", RSTITEMMAST!TELNO)
                txtfaxno.Text = IIf(IsNull(RSTITEMMAST!FAXNO), "", RSTITEMMAST!FAXNO)
                txtemail.Text = IIf(IsNull(RSTITEMMAST!EMAIL_ADD), "", RSTITEMMAST!EMAIL_ADD)
                TXTREMARKS.Text = IIf(IsNull(RSTITEMMAST!Remarks), "", RSTITEMMAST!Remarks)
                CmdDelete.Enabled = True
            End If
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
            Call CmdExit_Click
    End Select
Exit Sub
eRRHAND:
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
    On Error GoTo eRRHAND
    If REPFLAG = True Then
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='321')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='321')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
        'RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE ACT_NAME Like '" & txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
        REPFLAG = False
    End If
    Set DataList2.RowSource = RSTREP
    DataList2.ListField = "ACT_NAME"
    DataList2.BoundColumn = "ACT_CODE"
    
    Exit Sub
eRRHAND:
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
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub txtsupplist_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
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
    End Select
End Sub

Private Sub txttelno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

