VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmUserMast 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Creation"
   ClientHeight    =   5565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "Frmuser.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   7425
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
      Left            =   1905
      MaxLength       =   10
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   4470
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
      Left            =   1905
      MaxLength       =   3
      TabIndex        =   8
      Top             =   420
      Width           =   1455
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00FFC0C0&
      Height          =   4575
      Left            =   0
      TabIndex        =   1
      Top             =   975
      Width           =   7395
      Begin VB.OptionButton OptSales2 
         BackColor       =   &H0000C000&
         Caption         =   "Sales Section (Limited)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   3900
         TabIndex        =   23
         Top             =   2775
         Width           =   1905
      End
      Begin VB.OptionButton optadmn2 
         BackColor       =   &H0000C000&
         Caption         =   "Admin Privilage but hide profit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   60
         TabIndex        =   19
         Top             =   3270
         Width           =   1875
      End
      Begin VB.OptionButton OptUser3 
         BackColor       =   &H0000C000&
         Caption         =   "Manager"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   3900
         TabIndex        =   22
         Top             =   3270
         Width           =   1905
      End
      Begin VB.OptionButton Optuser2 
         BackColor       =   &H0000C000&
         Caption         =   "Purchase Section"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   450
         Left            =   1965
         TabIndex        =   21
         Top             =   3270
         Width           =   1905
      End
      Begin VB.OptionButton OptAdmin 
         BackColor       =   &H0000C000&
         Caption         =   "Administrator"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   60
         TabIndex        =   18
         Top             =   2775
         Width           =   1875
      End
      Begin VB.OptionButton OptUser 
         BackColor       =   &H0000C000&
         Caption         =   "Sales Section"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   1965
         TabIndex        =   20
         Top             =   2775
         Width           =   1905
      End
      Begin VB.TextBox Txtoldpass 
         Appearance      =   0  'Flat
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
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2190
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   16
         Top             =   870
         Width           =   5160
      End
      Begin VB.TextBox Txtpassword2 
         Appearance      =   0  'Flat
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
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2175
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   2055
         Width           =   5175
      End
      Begin VB.TextBox Txtpassword1 
         Appearance      =   0  'Flat
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
         Height          =   390
         IMEMode         =   3  'DISABLE
         Left            =   2175
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   12
         Top             =   1440
         Width           =   5175
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
         Left            =   4890
         MaskColor       =   &H80000007&
         TabIndex        =   6
         Top             =   3885
         UseMaskColor    =   -1  'True
         Width           =   1275
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
         Left            =   3480
         MaskColor       =   &H80000007&
         TabIndex        =   5
         Top             =   3885
         UseMaskColor    =   -1  'True
         Width           =   1305
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
         Left            =   2070
         MaskColor       =   &H80000007&
         TabIndex        =   4
         Top             =   3885
         UseMaskColor    =   -1  'True
         Width           =   1305
      End
      Begin VB.TextBox txtuser 
         Appearance      =   0  'Flat
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
         Height          =   390
         Left            =   2205
         MaxLength       =   50
         TabIndex        =   3
         Top             =   285
         Width           =   5145
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
         Left            =   195
         MaskColor       =   &H80000007&
         TabIndex        =   2
         Top             =   3885
         UseMaskColor    =   -1  'True
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Old Password"
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
         Height          =   300
         Index           =   6
         Left            =   105
         TabIndex        =   17
         Top             =   930
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Re-type Password"
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
         Height          =   300
         Index           =   5
         Left            =   90
         TabIndex        =   15
         Top             =   2115
         Width           =   2040
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
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
         Height          =   300
         Index           =   4
         Left            =   90
         TabIndex        =   13
         Top             =   1515
         Width           =   1605
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
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
         Height          =   300
         Index           =   1
         Left            =   90
         TabIndex        =   7
         Top             =   330
         Width           =   1605
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1905
      TabIndex        =   9
      Top             =   780
      Visible         =   0   'False
      Width           =   4470
      _ExtentX        =   7885
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
      Caption         =   "User Code"
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
      Left            =   45
      TabIndex        =   11
      Top             =   420
      Width           =   1920
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
      ForeColor       =   &H00000040&
      Height          =   300
      Index           =   8
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6300
   End
End
Attribute VB_Name = "frmUserMast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTREP As New ADODB.Recordset
Dim REPFLAG As Boolean
Dim NEW_USER As Boolean

Private Sub cmdcancel_Click()
    FRAME.Visible = False
    txtuser.text = ""
    Txtpassword1.text = ""
    Txtpassword2.text = ""
    Txtpassword1.text = ""
    Txtoldpass.text = ""
    Txtpassword2.text = ""
    Txtsuplcode.Enabled = True
    OptUser.Value = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * From USERS WHERE USER_ID = '" & Txtsuplcode.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        MsgBox "User doesn't exists...", vbOKOnly, "USER CREATION"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        txtuser.SetFocus
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * From USERS", db, adOpenStatic, adLockReadOnly, adCmdText
    If RSTITEMMAST.RecordCount = 1 Then
        MsgBox "Atleast one user must exists...", vbOKOnly, "USER CREATION"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        txtuser.SetFocus
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * From USERS WHERE LEVEL = '0'", db, adOpenStatic, adLockReadOnly, adCmdText
    If RSTITEMMAST.RecordCount = 1 Then
        If RSTITEMMAST!USER_ID = Txtsuplcode.text Then
            MsgBox "Atleast one user must be Administrator...", vbOKOnly, "USER CREATION"
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            txtuser.SetFocus
            Exit Sub
        End If
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & txtuser.text & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    db.Execute "delete  From USERS WHERE USER_ID = '" & Txtsuplcode.text & "' "
    MsgBox "USER DELETED SUCCESSFULLY..", vbOKOnly, "USER_NAME CREATION"
    FRAME.Visible = False
    txtuser.text = ""
    Txtpassword1.text = ""
    Txtpassword2.text = ""
    Txtpassword1.text = ""
    Txtoldpass.text = ""
    Txtpassword2.text = ""
    Txtsuplcode.Enabled = True
    OptUser.Value = True
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    Dim Request As String, NewID As Long, rsData As Recordset
    Dim MD5 As New clsMD5, NewPassword As String, OldPassword As String
    
    If Trim(txtuser.text) = "" Then
        MsgBox "ENTER NAME OF THE USER", vbOKOnly, "USER CREATION"
        txtuser.SetFocus
        Exit Sub
    End If
    
    If OptAdmin.Value = False Then
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * From USERS WHERE LEVEL = '0'", db, adOpenStatic, adLockReadOnly, adCmdText
        If RSTITEMMAST.RecordCount = 1 Then
            If RSTITEMMAST!USER_ID = Txtsuplcode.text Then
                MsgBox "Atleast one user must be Administrator...", vbOKOnly, "USER CREATION"
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                txtuser.SetFocus
                Exit Sub
            End If
        End If
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    
    NewPassword = MD5.DigestStrToHexStr(Me.Txtpassword1.text)
    OldPassword = MD5.DigestStrToHexStr(Me.Txtoldpass.text)
    
    Set rsData = db.Execute("SELECT PASS_WORD FROM USERS WHERE USER_ID = '" & Txtsuplcode.text & "'")
    'Set RSTITEMMAST = New ADODB.Recordset
    'RSTITEMMAST.Open "SELECT * From USERS WHERE USER_ID = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If Not (rsData.BOF Or rsData.EOF) Then
        If OldPassword <> rsData("PASS_WORD").Value Then
            MsgBox "You have entered a wrong password", , "User Creation"
            Txtoldpass.Enabled = True
            Txtoldpass.SetFocus
            Exit Sub
         Else
            If Trim(Txtoldpass.text) <> "" Then Txtoldpass.text = ""
        End If
    End If
            
    If Txtpassword1.text = "" Then
        MsgBox "Please enter a password", vbOKOnly, "USER CREATION"
        Txtpassword1.SetFocus
        Exit Sub
    End If
    
    If Len(Txtpassword1.text) < 3 Then
        MsgBox "Password must contains atleast 3 characters", vbOKOnly, "USER CREATION"
        Txtpassword1.SetFocus
        Exit Sub
    End If
    
    If Txtpassword2.text <> Txtpassword1.text Then
        MsgBox "Password doesn't match", vbOKOnly, "USER CREATION"
        Txtpassword2.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * From USERS WHERE USER_NAME = '" & Trim(txtuser.text) & "' and USER_ID <> '" & Txtsuplcode.text & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If RSTITEMMAST.RecordCount > 0 Then
        MsgBox "The Name already exists", vbOKOnly, "USER CREATION"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        txtuser.SetFocus
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * From USERS WHERE USER_ID = '" & Txtsuplcode.text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!USER_ID = Val(Txtsuplcode.text)
    End If
    RSTITEMMAST!USER_NAME = Trim(txtuser.text)
    RSTITEMMAST!PASS_WORD = NewPassword
    If OptAdmin.Value = True Then
        RSTITEMMAST!Level = "0"
    ElseIf OptUser.Value = True Then
        RSTITEMMAST!Level = "2"
    ElseIf Optuser2.Value = True Then
        RSTITEMMAST!Level = "3"
    ElseIf OptUser3.Value = True Then
        RSTITEMMAST!Level = "1"
    ElseIf optadmn2.Value = True Then
        RSTITEMMAST!Level = "4"
    ElseIf OptSales2.Value = True Then
        RSTITEMMAST!Level = "5"
    Else
        RSTITEMMAST!Level = "0"
    End If
    
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "USER_NAME CREATION"
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(CONVERT(USER_ID, SIGNED INTEGER)) From USERS ", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        Txtsuplcode.text = IIf(IsNull(RSTITEMMAST.Fields(0)), 1, RSTITEMMAST.Fields(0) + 1)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    FRAME.Visible = False
    txtuser.text = ""
    Txtpassword1.text = ""
    Txtpassword2.text = ""
    Txtpassword1.text = ""
    Txtoldpass.text = ""
    Txtpassword2.text = ""
    Txtsuplcode.Enabled = True
    OptUser.Value = True
Exit Sub
ErrHand:
    MsgBox (err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Txtpassword2.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ErrHand
            If Trim(DataList2.BoundText) = "" Then Exit Sub
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT USER_ID From USERS WHERE USER_ID = '" & Trim(DataList2.BoundText) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.text = RSTITEMMAST!USER_ID
            End If
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox (err.Description)
End Sub

Private Sub Form_Activate()
    If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    NEW_USER = True
    'TMPFLAG = True
    REPFLAG = True
    FRAME.Visible = False
    'txtunit.Visible = False
    On Error GoTo ErrHand
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(USER_ID, SIGNED INTEGER)) From USERS ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Me.Top = 1000
    Me.Left = 1000
    Me.Height = 7065
    Me.Width = 7515
    Exit Sub
ErrHand:
    MsgBox (err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 555
    MDIMAIN.PCTMENU.SetFocus
    If REPFLAG = False Then RSTREP.Close
End Sub

Private Sub Txtoldpass_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim RSTITEMMAST As ADODB.Recordset, Old_Password As String, MD5 As New clsMD5

    Old_Password = MD5.DigestStrToHexStr(Txtoldpass.text)
    
    Select Case KeyCode
        Case vbKeyReturn
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * From USERS WHERE USER_ID = '" & Txtsuplcode.text & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                If RSTITEMMAST!PASS_WORD <> Old_Password Then
                    MsgBox "You have entered a wrong password", , "User Creation"
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    Txtoldpass.SetFocus
                    Exit Sub
                End If
                Txtpassword1.SetFocus
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing

            Txtpassword1.SetFocus
        Case vbKeyEscape
            txtuser.SetFocus
    End Select
    
End Sub

Private Sub Txtoldpass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub txtuser_GotFocus()
    txtuser.SelStart = 0
    txtuser.SelLength = Len(txtuser.text)
   
End Sub

Private Sub txtuser_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtuser.text) = "" Then
                MsgBox "ENTER THE NAME FOR OFFICE", vbOKOnly, "USER CREATION"
                txtuser.SetFocus
                Exit Sub
            End If
            If NEW_USER = False Then
                Txtoldpass.Enabled = True
                Txtoldpass.SetFocus
            Else
                Txtoldpass.Enabled = False
                Txtpassword1.SetFocus
            End If
    End Select
    
End Sub

Private Sub txtuser_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc(",")
'            KeyAscii = Asc(Chr(KeyAscii))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub Txtsuplcode_GotFocus()
    Txtsuplcode.SelStart = 0
    Txtsuplcode.SelLength = Len(Txtsuplcode.text)
End Sub

Private Sub Txtsuplcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Txtsuplcode.text) = "" Then Exit Sub
            'If Val(Txtsuplcode.Text) = 1 Then Exit Sub
            On Error GoTo ErrHand
            
            NEW_USER = True
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "Select MAX(CONVERT(USER_ID, SIGNED INTEGER)) From USERS ", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                i = IIf(IsNull(RSTITEMMAST.Fields(0)), 0, RSTITEMMAST.Fields(0))
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            If Val(Txtsuplcode.text) > i Then Txtsuplcode.text = i + 1
            
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * From USERS WHERE USER_ID = '" & Txtsuplcode.text & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                NEW_USER = False
                If RSTITEMMAST!Level = "0" Then
                    OptAdmin.Value = True
                ElseIf RSTITEMMAST!Level = "1" Then
                    OptUser3.Value = True
                ElseIf RSTITEMMAST!Level = "2" Then
                    OptUser.Value = True
                ElseIf RSTITEMMAST!Level = "3" Then
                    Optuser2.Value = True
                ElseIf RSTITEMMAST!Level = "4" Then
                    optadmn2.Value = True
                ElseIf RSTITEMMAST!Level = "5" Then
                    OptSales2.Value = True
                End If
                txtuser.text = RSTITEMMAST!USER_NAME
                CmdDelete.Enabled = True
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Txtsuplcode.Enabled = False
            FRAME.Visible = True
            txtuser.SetFocus
        Case 114
            txtsupplist.text = ""
            txtsupplist.Visible = True
            DataList2.Visible = True
            txtsupplist.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
Exit Sub
ErrHand:
    MsgBox err.Description
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
        RSTREP.Open "Select USER_ID, USER_NAME From USERS WHERE USER_NAME Like '" & Me.txtsupplist.text & "%'ORDER BY USER_ID", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select USER_ID, USER_NAME From USERS WHERE USER_NAME Like '" & Me.txtsupplist.text & "%'ORDER BY USER_ID", db, adOpenStatic, adLockReadOnly
        'RSTREP.Open "Select USER_ID,USER_NAME From USERS  WHERE USER_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY USER_NAME", db, adOpenStatic,adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "USER_NAME"
    DataList2.BoundColumn = "USER_ID"
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub txtsupplist_GotFocus()
    txtsupplist.SelStart = 0
    txtsupplist.SelLength = Len(txtsupplist.text)
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
    MsgBox err.Description
    
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

Private Sub Txtpassword1_GotFocus()
    Txtpassword1.SelStart = 0
    Txtpassword1.SelLength = Len(Txtpassword1.text)
End Sub

Private Sub Txtpassword1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Txtpassword1.text) = "" Then
                MsgBox "ENTER THE DISPLAY NAME", vbOKOnly, "USER CREATION"
                Txtpassword1.SetFocus
                Exit Sub
            End If
            If Len(Txtpassword1.text) < 3 Then
                MsgBox "Password must contains atleast 3 characters", vbOKOnly, "USER CREATION"
                Txtpassword1.SetFocus
                Exit Sub
            End If
            Txtpassword2.SetFocus
        Case vbKeyEscape
            If NEW_USER = True Then
                txtuser.SetFocus
            Else
                Txtoldpass.SetFocus
            End If
    End Select
    
End Sub

Private Sub Txtpassword1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpassword2_GotFocus()
    Txtpassword2.SelStart = 0
    Txtpassword2.SelLength = Len(Txtpassword2.text)
End Sub

Private Sub Txtpassword2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = vbCtrlMask Then Txtpassword2.text = Clipboard.GetText
    Select Case KeyCode
        Case vbKeyReturn
            If Txtpassword2.text <> Txtpassword1.text Then
                MsgBox "Password doesn't match", vbOKOnly, "USER CREATION"
                Txtpassword2.SetFocus
                Exit Sub
            End If
            cmdSAVE.SetFocus
        Case vbKeyEscape
            Txtpassword1.SetFocus
    End Select
End Sub

Private Sub Txtpassword2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

