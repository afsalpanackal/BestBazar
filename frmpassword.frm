VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FFC0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   4740
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   5775
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   8.25
      Charset         =   0
      Weight          =   900
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmpassword.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   5775
   Begin VB.TextBox txtretype 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1965
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2700
      Width           =   3675
   End
   Begin VB.TextBox txtnewpass 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1965
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2115
      Width           =   3675
   End
   Begin VB.TextBox txtoldpass 
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
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1965
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1560
      Width           =   3675
   End
   Begin VB.CommandButton cmdEXIT 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4335
      TabIndex        =   5
      Top             =   3300
      Width           =   1230
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2895
      TabIndex        =   4
      Top             =   3315
      Width           =   1320
   End
   Begin VB.TextBox txtPassword 
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
      Height          =   405
      Left            =   1980
      TabIndex        =   6
      Top             =   3930
      Visible         =   0   'False
      Width           =   3675
   End
   Begin VB.TextBox txtLogin 
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
      Height          =   375
      Left            =   1965
      TabIndex        =   0
      Top             =   990
      Width           =   3675
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Retype Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   225
      TabIndex        =   11
      Top             =   2760
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   225
      TabIndex        =   10
      Top             =   1605
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "User Credentials"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   345
      Left            =   240
      TabIndex        =   9
      Top             =   345
      Width           =   5070
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Login:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   255
      TabIndex        =   8
      Top             =   990
      Width           =   1065
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   225
      TabIndex        =   7
      Top             =   2220
      Width           =   1575
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsPass As ADODB.Recordset

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
         Case vbKeyEscape
            cmdUpdate.Enabled = False
            txtretype.Enabled = True
            txtretype.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
'Place form in center

End Sub

Private Sub Form_Load()

Dim sql As String

Set rsPass = New ADODB.Recordset
sql = "select * from passwords where Target ='this'"
rsPass.Open sql, db2, adOpenKeyset, adLockPessimistic

If Not (rsPass.EOF Or rsPass.BOF) Then
    txtLogin.Text = rsPass!Login
    txtPassword.Text = rsPass!Password
End If
Me.Left = 0
Me.Top = 0
End Sub

Private Sub cmdUpdate_Click()

Dim sql As String
Dim target As String


'If any of the field is empty warning message displays
If txtLogin.Text = "" Or txtnewpass.Text = "" Then
MsgBox "Please enter login and password", vbOKOnly, "Update"
Exit Sub
End If

target = "this"

'If all fields have data load field values to recordset

rsPass!Login = txtLogin.Text
rsPass!Password = txtnewpass.Text

'If user did not enter a new target in target text field the use combobox selected value.
'Or use new target entered in txtTarget field.
rsPass!target = target

rsPass!note = ""
rsPass.Update
txtPassword.Text = txtnewpass.Text
txtoldpass.Text = txtnewpass.Text
txtnewpass.Text = ""
txtretype.Text = ""
txtnewpass.Enabled = False
txtretype.Enabled = False
cmdUpdate.Enabled = False
MsgBox "Saved Successfully", vbOKOnly, "Update"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MDIMAIN.PCTMENU.Visible = True Then
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    Else
        MDIMAIN.pctmenu2.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.pctmenu2.SetFocus
    End If
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
End Sub

Private Sub txtLogin_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If txtLogin.Text = "" Then Exit Sub
            txtoldpass.SetFocus
    End Select
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtnewpass_GotFocus()
    txtnewpass.SelStart = 0
    txtnewpass.SelLength = Len(txtnewpass.Text)
End Sub

Private Sub txtnewpass_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtnewpass.Text) = "" Then Exit Sub
            txtnewpass.Enabled = False
            txtretype.Enabled = True
            txtretype.SetFocus
    End Select
End Sub

Private Sub txtnewpass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtoldpass_GotFocus()
    txtoldpass.SelStart = 0
    txtoldpass.SelLength = Len(txtoldpass.Text)
End Sub

Private Sub txtoldpass_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtoldpass.Text = txtPassword.Text Then
                txtnewpass.Enabled = True
                txtnewpass.SetFocus
            Else
                MsgBox "Password Incorrect", vbOKOnly, "Login"
            End If
    End Select
End Sub

Private Sub txtoldpass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtPassword.Text = "" Then Exit Sub
            cmdUpdate.SetFocus
    End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtretype_GotFocus()
    txtretype.SelStart = 0
    txtretype.SelLength = Len(txtretype.Text)
End Sub

Private Sub txtretype_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If txtnewpass.Text = txtretype.Text Then
                txtretype.Enabled = False
                cmdUpdate.Enabled = True
                cmdUpdate.SetFocus
            Else
                MsgBox "The password you type do not match", vbOKOnly, "Login...."
            End If
        Case vbKeyEscape
            txtretype.Enabled = False
            txtnewpass.Enabled = True
            txtnewpass.SetFocus
    End Select
End Sub

Private Sub txtretype_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
