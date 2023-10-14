VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMMAIL1 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3750
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   5580
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3750
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   3690
      Left            =   45
      TabIndex        =   9
      Top             =   45
      Width           =   5460
      Begin VB.CommandButton CmdRemove 
         Caption         =   "&Remove Attachment"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         TabIndex        =   5
         Top             =   3120
         Width           =   1710
      End
      Begin VB.CommandButton CmdAttach 
         Caption         =   "&Attach File"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1875
         TabIndex        =   6
         Top             =   3120
         Width           =   1125
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   3045
         TabIndex        =   7
         Top             =   3120
         Width           =   1110
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "E-mail (Yahoo Mail)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   3495
         Left            =   60
         TabIndex        =   10
         Top             =   150
         Width           =   5340
         Begin VB.TextBox TXTSENDER 
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
            IMEMode         =   3  'DISABLE
            Left            =   1335
            TabIndex        =   1
            Top             =   1230
            Width           =   3840
         End
         Begin VB.CheckBox chkpassword 
            BackColor       =   &H00FFC0C0&
            Caption         =   "&Save Password"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1335
            TabIndex        =   3
            Top             =   2070
            Value           =   1  'Checked
            Width           =   2025
         End
         Begin VB.TextBox txtpassword 
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
            IMEMode         =   3  'DISABLE
            Left            =   1350
            PasswordChar    =   "*"
            TabIndex        =   2
            Top             =   1665
            Width           =   3840
         End
         Begin VB.CheckBox chkcopy 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Send a &copy to own mail address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   90
            TabIndex        =   4
            Top             =   2640
            Value           =   1  'Checked
            Width           =   3465
         End
         Begin VB.CommandButton CMDEXIT 
            Caption         =   "&Exit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   4140
            TabIndex        =   8
            Top             =   2970
            Width           =   1110
         End
         Begin VB.Label Lblform 
            Height          =   180
            Left            =   3615
            TabIndex        =   16
            Top             =   2610
            Visible         =   0   'False
            Width           =   1560
         End
         Begin VB.Label lblpath 
            Height          =   210
            Left            =   3615
            TabIndex        =   15
            Top             =   2310
            Visible         =   0   'False
            Width           =   1470
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Sender"
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
            Left            =   75
            TabIndex        =   14
            Top             =   1260
            Width           =   1110
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password"
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
            Left            =   90
            TabIndex        =   13
            Top             =   1695
            Width           =   1110
         End
         Begin MSForms.ComboBox txtmail 
            Height          =   375
            Left            =   1350
            TabIndex        =   0
            Top             =   345
            Width           =   3840
            VariousPropertyBits=   746604571
            ForeColor       =   255
            MaxLength       =   35
            DisplayStyle    =   3
            Size            =   "6773;661"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            DropButtonStyle =   0
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.Label lblattach 
            Height          =   270
            Left            =   1305
            TabIndex        =   12
            Top             =   2340
            Width           =   3855
            ForeColor       =   192
            BackColor       =   16761024
            Size            =   "6800;476"
            BorderColor     =   16761024
            FontName        =   "Times New Roman"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Mail address"
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
            Index           =   41
            Left            =   60
            TabIndex        =   11
            Top             =   390
            Width           =   1350
         End
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FRMMAIL1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private attach3 As String

Private Sub CmdAttach_Click()
    Dim i As Long
    Dim Date_flag As Boolean
    Dim rstAttach As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo errhandler
    lblattach.Caption = ""
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "All Files (*.*)|*.*"
    CommonDialog1.ShowOpen
    attach3 = CommonDialog1.FileName
    lblattach.Caption = "File " & CommonDialog1.FileName & " attached"
    lblpath.Caption = CommonDialog1.FileName
    Exit Sub
errhandler:
    Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select

End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
   
'    Dim RSTMAIL As ADODB.Recordset
'    Dim mail1 As String
'
'    On Error GoTo errhand
'
'    If Trim(TXTSENDER.Text) = "" Then
'        MsgBox "No Sender Email address available", , "E-Mail..."
'        TXTSENDER.SetFocus
'        Exit Sub
'    End If
'
'    If InStr(Trim(TXTSENDER.Text), "@") = 0 Or _
'        InStr(Trim(TXTSENDER.Text), ".") = 0 Or _
'        Len(Trim(TXTSENDER.Text)) < 7 Then
'        MsgBox "Invalid Sender Email Address", , "E-Mail..."
'        TXTSENDER.SetFocus
'        Exit Sub
'    End If
'
'    If Trim(txtPassword.Text) = "" Then
'        MsgBox "Please enter password for sender maill id", , "E-Mail..."
'        txtPassword.SetFocus
'        Exit Sub
'    End If
'
'    Set RSTMAIL = New ADODB.Recordset
'    RSTMAIL.Open "SELECT * from PASSWORD ", db, adOpenStatic, adLockOptimistic, adCmdText
'    If (RSTMAIL.EOF And RSTMAIL.BOF) Then
'        RSTMAIL.AddNew
'    End If
'    If chkpassword.value = 1 Then
'        RSTMAIL!Password = txtPassword.Text
'    Else
'        RSTMAIL!Password = ""
'    End If
'    RSTMAIL!MAIL_ID = Trim(TXTSENDER.Text)
'    RSTMAIL.Update
'
'    RSTMAIL.Close
'    Set RSTMAIL = Nothing
'
'    If Trim(txtmail.Text) = "" Then
'        MsgBox "No Receipoent Email address available", , "E-Mail..."
'        txtmail.SetFocus
'        Exit Sub
'    End If
'
'    If InStr(Trim(txtmail.Text), "@") = 0 Or _
'        InStr(Trim(txtmail.Text), ".") = 0 Or _
'        Len(Trim(txtmail.Text)) < 7 Then
'        MsgBox "Invalid Receipient Email Address", , "E-Mail..."
'        txtmail.SetFocus
'        Exit Sub
'    End If
'
'    If lblpath.Caption = "" Then
'    'If Dir("E:\MailOUT\" & "Bill No" & frmlw.txtBillNo.Text & ".pdf", vbDirectory) = "" Then
'        MsgBox "File Not exists... Please Re-print....", , "E-mail"
'        Exit Sub
'    End If
'    If MsgBox("This will Send " & Lblform.Caption & " Bill to " & Trim(txtmail.Text) & Chr(13) & "Are you sure...?", vbYesNo, "E-Mail") = vbNo Then Exit Sub
'
'    Screen.MousePointer = vbHourglass
'    Dim oSmtp As New EASendMailObjLib.Mail
'    oSmtp.LicenseCode = "TryIt"
'
'    ' Set your Yahoo email address
'    oSmtp.FromAddr = Trim(TXTSENDER.Text)
'
'    ' Add recipient email address
'    oSmtp.AddRecipientEx Trim(txtmail.Text), 0
'    If chkcopy.value = 1 Then
'        oSmtp.AddRecipientEx Trim(TXTSENDER.Text), 0
'    End If
'    oSmtp.Subject = Lblform.Caption & " Bill"
'    oSmtp.BodyText = "Please see the attached " & Lblform.Caption & " Bill." & "<br>" & "Thanks & Regards.. " & "<br>" & ""
'
'    Dim attach1, attach2 As String
'    attach1 = lblpath.Caption
'    If oSmtp.AddAttachment(attach1) <> 0 Then
'        MsgBox oSmtp.GetLastErrDescription() & ":" & attach1
'        Screen.MousePointer = vbNormal
'        'btnSend.Enabled = True
'        Exit Sub
'    End If
'
''    If oSmtp.AddAttachment(attach2) <> 0 Then
''        MsgBox oSmtp.GetLastErrDescription() & ":" & attach2
''        Screen.MousePointer = vbNormal
''        'btnSend.Enabled = True
''        Exit Sub
''    End If
''
'    If attach3 <> "" Then
'        If oSmtp.AddAttachment(attach3) <> 0 Then
'            MsgBox oSmtp.GetLastErrDescription() & ":" & attach3
'            Screen.MousePointer = vbNormal
'            'btnSend.Enabled = True
'            Exit Sub
'        End If
'    End If
'
'    If Trim(txtmail.Text) <> "" Then
'        Set RSTMAIL = New ADODB.Recordset
'        RSTMAIL.Open "SELECT DISTINCT EMAIL_ID from ADDRESS_BOOK WHERE EMAIL_ID = '" & Trim(txtmail.Text) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
'        If (RSTMAIL.EOF And RSTMAIL.BOF) Then
'            RSTMAIL.AddNew
'            RSTMAIL!EMAIL_ID = Trim(txtmail.Text)
'            RSTMAIL.Update
'        End If
'        RSTMAIL.Close
'        Set RSTMAIL = Nothing
'    End If
'
''    oSmtp.ServerAddr = "smtp.gmail.com"
'
'
'    oSmtp.ServerAddr = "smtp.mail.yahoo.com"
'    oSmtp.UserName = Trim(TXTSENDER.Text)
'    oSmtp.Password = txtPassword.Text
'    oSmtp.ServerPort = 587
'    oSmtp.SSL_starttls = 1
'
'    'oSmtp.ServerPort = 465
'    oSmtp.SSL_starttls = True
'
'    ' Detect SSL/TLS automatically
'    oSmtp.SSL_init
'
'    If oSmtp.SendMail() = 0 Then
'        Screen.MousePointer = vbNormal
'        MsgBox "Email was sent successfully!", , "E-Mail"
'        'creditbill.Enabled = True
'        Unload Me
'    Else
'        Screen.MousePointer = vbNormal
'        MsgBox "Failed to send email with the following error:" & oSmtp.GetLastErrDescription(), , "E-Mail"
'    End If
'    Screen.MousePointer = vbNormal
'    Exit Sub
'
'errhand:
'    Screen.MousePointer = vbNormal
'     MsgBox Err.Description
End Sub

Private Sub CmdRemove_Click()
    attach3 = ""
    lblattach.Caption = ""
End Sub

Private Sub Form_Load()
    Dim RSTMAIL As ADODB.Recordset
    
    On Error GoTo errhand
    Set RSTMAIL = New ADODB.Recordset
    RSTMAIL.Open "Select DISTINCT EMAIL_ID From ADDRESS_BOOK ORDER BY EMAIL_ID", db, adOpenForwardOnly
    Do Until RSTMAIL.EOF
        If Not IsNull(RSTMAIL!EMAIL_ID) Then txtmail.AddItem (RSTMAIL!EMAIL_ID)
        RSTMAIL.MoveNext
    Loop
    RSTMAIL.Close
    Set RSTMAIL = Nothing
    
    Set RSTMAIL = New ADODB.Recordset
    RSTMAIL.Open "SELECT * from PASSWORDS ", db, adOpenStatic, adLockReadOnly
    If Not (RSTMAIL.EOF And RSTMAIL.BOF) Then
        txtPassword.Text = IIf(IsNull(RSTMAIL!PASSWORDS), "", RSTMAIL!PASSWORDS)
        TXTSENDER.Text = IIf(IsNull(RSTMAIL!MAIL_ID), "", RSTMAIL!MAIL_ID)
    End If
    RSTMAIL.Close
    Set RSTMAIL = Nothing
    
    cetre Me
    MDIMAIN.Enabled = False
    Exit Sub
errhand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'creditbill.Enabled = True
    MDIMAIN.Enabled = True
End Sub

Private Sub TXTMAIL_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    On Error GoTo errhand
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK.SetFocus
        Case vbKeyEscape
            Unload Me
    End Select
    Exit Sub

errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTMAIL_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtPassword.Text) = "" Then Exit Sub
            cmdOK.SetFocus
        Case vbKeyEscape
            TXTSENDER.SetFocus
    End Select
End Sub

Private Sub TXTSENDER_GotFocus()
    TXTSENDER.SelStart = 0
    TXTSENDER.SelLength = Len(TXTSENDER.Text)
End Sub

Private Sub TXTSENDER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTSENDER.Text) = "" Then Exit Sub
            txtPassword.SetFocus
        Case vbKeyEscape
            txtmail.SetFocus
    End Select
End Sub
