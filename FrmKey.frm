VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmKey 
   Caption         =   "Key Activation"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5220
   Icon            =   "FrmKey.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4920
   ScaleWidth      =   5220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Continue"
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
      Left            =   135
      TabIndex        =   12
      Top             =   3555
      Width           =   1245
   End
   Begin VB.TextBox txtRegnKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   9
      Top             =   2850
      Width           =   4905
   End
   Begin VB.TextBox txtregn 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   150
      TabIndex        =   7
      Top             =   2025
      Width           =   4905
   End
   Begin VB.TextBox TxtKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   3
      Top             =   1245
      Width           =   4905
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
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
      Left            =   3795
      TabIndex        =   1
      Top             =   3555
      Width           =   1245
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "&Load"
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
      Left            =   2385
      TabIndex        =   0
      Top             =   3555
      Width           =   1245
   End
   Begin MSMask.MaskEdBox txtRegnKey2 
      Height          =   375
      Left            =   150
      TabIndex        =   11
      Top             =   4740
      Visible         =   0   'False
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   24
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "####-####-####-####-####"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      Caption         =   "Registration Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   165
      TabIndex        =   10
      Top             =   2565
      Width           =   1845
   End
   Begin VB.Label Label3 
      Caption         =   "Registration ID"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   165
      TabIndex        =   8
      Top             =   1740
      Width           =   1845
   End
   Begin VB.Label lblMsg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   165
      TabIndex        =   6
      Top             =   3210
      Width           =   4875
   End
   Begin VB.Label lblINSID 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   150
      TabIndex        =   5
      Top             =   540
      Width           =   4890
   End
   Begin VB.Label Label2 
      Caption         =   "Activation Key"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   165
      TabIndex        =   4
      Top             =   975
      Width           =   1845
   End
   Begin VB.Label Label1 
      Caption         =   "Installation ID:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   270
      Left            =   165
      TabIndex        =   2
      Top             =   270
      Width           =   1845
   End
End
Attribute VB_Name = "FrmKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)
Option Explicit

Private Sub CmdExit_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
    
    If IsFormLoaded(MDIMAIN) = True Then
        Screen.MousePointer = vbNormal
        MDIMAIN.Enabled = True
        Exit Sub
    End If
    On Error Resume Next
    frmLogin.rs!LAST_LOGOUT = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss")
    frmLogin.rs.Update
    frmLogin.rs.Close
    Set frmLogin.rs = Nothing
    db.Close
    Set db = Nothing
    Screen.MousePointer = vbNormal
    End
End Sub

Private Sub CmdLoad_Click()
    Dim sql As String
    Dim MD5 As New clsMD5
    Dim ACT_KEY1, ACT_KEY2 As String
    Dim TRXFILE As ADODB.Recordset
    
    On Error GoTo ERRHAND
    ACT_KEY1 = lblINSID.Caption
    ACT_KEY2 = UCase(MD5.DigestStrToHexStr(ACT_KEY1))
    ACT_KEY2 = ACT_KEY2 & UCase(MD5.DigestStrToHexStr(ACT_KEY2))
    ACT_KEY2 = Mid(ACT_KEY2, 24, 10) & Mid(ACT_KEY2, 1, 5)
    
    If ACT_KEY2 <> Trim(TxtKey.text) Then
        MsgBox "Invalid Key", vbOKOnly, "Product Activation"
        Exit Sub
    End If
    
'    Set TRXFILE = New ADODB.Recordset
'    TRXFILE.Open "Select MAX(COMP_CODE) From COMPINFO", db, adOpenStatic, adLockReadOnly
'    If Not (TRXFILE.EOF And TRXFILE.BOF) Then
'        If IsNull(TRXFILE.Fields(0)) Then
'            Max_Com_Code = 1
'        Else
'            Max_Com_Code = Val(TRXFILE.Fields(0)) + 1
'        End If
'    End If
'    TRXFILE.Close
'    Set TRXFILE = Nothing
    
   
    Set TRXFILE = New ADODB.Recordset
    'sql = "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "' "
    sql = "select * from act_ky WHERE ACT_CODE = '" & ACT_KEY2 & "' "
    TRXFILE.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
    If (TRXFILE.BOF And TRXFILE.EOF) Then
        TRXFILE.AddNew
        'LblRegnCode.Text = Mid(txtRegnKey.Text, 1, 4) & "-" & Mid(txtRegnKey.Text, 5, 4) & "-" & Mid(txtRegnKey.Text, 9, 4) & "-" & Mid(txtRegnKey.Text, 13, 4) & "-" & Mid(Text.Caption, 17, 4)
        TRXFILE!actky2 = Trim(txtregn.text)
        TRXFILE!actky3 = Left(Trim(txtRegnKey.text), 20)
        ''TRXFILE!actky4 = EncryptString(Mid(Trim(txtRegnKey.Text), 21), "ezkeys")
        'TRXFILE!actky3 = Mid(txtRegnKey.Text, 1, 4) & Mid(txtRegnKey.Text, 5, 4) & Mid(txtRegnKey.Text, 9, 4) & Mid(txtRegnKey.Text, 13, 4) & Mid(Text.Caption, 17, 4)
        TRXFILE!ACT_CODE = ACT_KEY2
        TRXFILE!ACT_DATE = Format(Date, "DD/MM/YYYY")
        TRXFILE.Update
    Else
        TRXFILE!actky2 = Trim(txtregn.text)
        'TRXFILE!actky3 = Mid(txtRegnKey.Text, 1, 4) & Mid(txtRegnKey.Text, 5, 4) & Mid(txtRegnKey.Text, 9, 4) & Mid(txtRegnKey.Text, 13, 4) & Mid(Text.Caption, 17, 4)
        TRXFILE!actky3 = Left(Trim(txtRegnKey.text), 20)
        
        ''TRXFILE!actky4 = EncryptString(Mid(Trim(txtRegnKey.Text), 21), "ezkeys")
        TRXFILE.Update
    End If
    TRXFILE.Close
    Set TRXFILE = Nothing
    
    Dim expkey As String
    If Val(Mid(Trim(txtRegnKey.text), 21)) = 0 Then
        expkey = EncryptString(Trim(txtregn.text), "ezbizkeys")
        If IsDate(expkey) Then
            expkey = DateDiff("d", Date, expkey)
        Else
            expkey = 365
        End If
        expkey = EncryptString(DateAdd("d", Val(expkey), Date), "ezkeys")
    Else
        expkey = EncryptString(DateAdd("d", Val(Mid(Trim(txtRegnKey.text), 21)), Date), "ezkeys")
    End If
    Set TRXFILE = New ADODB.Recordset
    sql = "select * from act_ky "
    TRXFILE.Open sql, db, adOpenKeyset, adLockOptimistic, adCmdText
    Do Until TRXFILE.EOF
        TRXFILE!actky4 = expkey
        TRXFILE.Update
        TRXFILE.MoveNext
    Loop
    TRXFILE.Close
    Set TRXFILE = Nothing
    
    'db.Execute "Update act_ky set actky4 = expkey"
    
    db.Execute "Update COMPINFO set EC =0"
    MsgBox "Product Activation Success. Please restart the program", vbOKOnly, "Product Activation"
    Unload Me
    End
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TxtInsID_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub Form_Activate()
    Dim MD5 As New clsMD5
    Dim TRXFILE As ADODB.Recordset
    Dim ACT_KEY1, ACT_KEY2 As String
    
    On Error GoTo ERRHAND
    ACT_KEY1 = lblINSID.Caption
    ACT_KEY2 = UCase(MD5.DigestStrToHexStr(ACT_KEY1))
    ACT_KEY2 = ACT_KEY2 & UCase(MD5.DigestStrToHexStr(ACT_KEY2))
    ACT_KEY2 = Mid(ACT_KEY2, 24, 10) & Mid(ACT_KEY2, 1, 5)
    
    Set TRXFILE = New ADODB.Recordset
    'sql = "select * from act_ky WHERE ACT_CODE = '" & ACT_KEY2 & "' "
    TRXFILE.Open "select * from act_ky WHERE ACT_CODE = '" & ACT_KEY2 & "' ", db, adOpenKeyset, adLockReadOnly, adCmdText
    If Not (TRXFILE.BOF And TRXFILE.EOF) Then
        TxtKey.text = TRXFILE!ACT_CODE
    End If
    TRXFILE.Close
    Set TRXFILE = Nothing
    
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz Activation"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MDIMAIN.Enabled = True
End Sub

Private Sub lblINSID_DblClick()
    If Trim(lblINSID.Caption) = "" Then Exit Sub
    Clipboard.Clear
    Clipboard.SetText (lblINSID.Caption)
    lblMsg.Caption = "Installation ID Copied to Clipboard.."
    'MsgBox "Installation ID Copied", vbOKOnly, "Product Activation"
End Sub

' encrypt a string using a password
'
' you must reapply the same function (and same password) on
' the encrypted string to obtain the original, non-encrypted string
'
' you get better, more secure results if you use a long password
' (e.g. 16 chars or longer). This routine works well only with ANSI strings.

Function EncryptString(ByVal text As String, ByVal Password As String) As String
    Dim passLen As Long
    Dim i As Long
    Dim passChr As Integer
    Dim passNdx As Long
    
    passLen = Len(Password)
    ' null passwords are invalid
    If passLen = 0 Then err.Raise 5
    
    ' move password chars into an array of Integers to speed up code
    ReDim passChars(0 To passLen - 1) As Integer
    CopyMemory passChars(0), ByVal StrPtr(Password), passLen * 2
    
    ' this simple algorithm XORs each character of the string
    ' with a character of the password, but also modifies the
    ' password while it goes, to hide obvious patterns in the
    ' result string
    For i = 1 To Len(text)
        ' get the next char in the password
        passChr = passChars(passNdx)
        ' encrypt one character in the string
        Mid$(text, i, 1) = Chr$(Asc(Mid$(text, i, 1)) Xor passChr)
        ' modify the character in the password (avoid overflow)
        passChars(passNdx) = (passChr + 17) And 255
        ' prepare to use next char in the password
        passNdx = (passNdx + 1) Mod passLen
    Next

    EncryptString = text
End Function



