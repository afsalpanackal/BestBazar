VERSION 5.00
Begin VB.Form Frmexpkey 
   Caption         =   "Error Correction "
   ClientHeight    =   1650
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "Frmexpkey.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1650
   ScaleWidth      =   7620
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRegnKey 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   390
      Width           =   7425
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
      Left            =   6300
      TabIndex        =   1
      Top             =   885
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
      Left            =   4890
      TabIndex        =   0
      Top             =   885
      Width           =   1245
   End
   Begin VB.Label Label4 
      Caption         =   "Run Query here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
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
      TabIndex        =   3
      Top             =   105
      Width           =   1845
   End
End
Attribute VB_Name = "Frmexpkey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)
Option Explicit

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdLoad_Click()
        
    Dim expkey As String
    txtRegnKey.Text = Trim(txtRegnKey.Text)
    If Len(txtRegnKey.Text) < 40 And Len(txtRegnKey.Text) > 45 Then
        MsgBox "Error in query", , "EzBiz"
        Exit Sub
    Else
        expkey = Right(txtRegnKey.Text, Len(txtRegnKey.Text) - 31)
    End If
    Screen.MousePointer = vbHourglass
    Sleep (5000)
    
    
    Dim TRXFILE As ADODB.Recordset
    Dim sql As String
    
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
    
    'db.Execute "update act_ky set actky4 = '" & expkey & "' "
    
    Screen.MousePointer = vbNormal
    MsgBox "Success. Please restart the program and try", vbOKOnly, "Error Correction"
    Unload Me
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    txtRegnKey.SetFocus
End Sub

Private Sub Form_Load()
'    Dim MD5 As New clsMD5
'    Dim TRXFILE As ADODB.Recordset
'    Dim ACT_KEY1, ACT_KEY2 As String
    
    On Error GoTo Errhand
'    ACT_KEY1 = lblINSID.Caption
'    ACT_KEY2 = UCase(MD5.DigestStrToHexStr(ACT_KEY1))
'    ACT_KEY2 = ACT_KEY2 & UCase(MD5.DigestStrToHexStr(ACT_KEY2))
'    ACT_KEY2 = Mid(ACT_KEY2, 24, 10) & Mid(ACT_KEY2, 1, 5)
'
'    Set TRXFILE = New ADODB.Recordset
'    'sql = "select * from act_ky WHERE ACT_CODE = '" & ACT_KEY2 & "' "
'    TRXFILE.Open "select * from act_ky WHERE ACT_CODE = '" & ACT_KEY2 & "' ", db, adOpenKeyset, adLockReadOnly, adCmdText
'    If Not (TRXFILE.BOF And TRXFILE.EOF) Then
'        TxtKey.Text = TRXFILE!ACT_CODE
'    End If
'    TRXFILE.Close
'    Set TRXFILE = Nothing
    cetre Me
    Exit Sub
Errhand:
    MsgBox Err.Description, , "EzBiz Error Correctin Module"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'MDIMAIN.Enabled = True
End Sub

' encrypt a string using a password
'
' you must reapply the same function (and same password) on
' the encrypted string to obtain the original, non-encrypted string
'
' you get better, more secure results if you use a long password
' (e.g. 16 chars or longer). This routine works well only with ANSI strings.

Function EncryptString(ByVal Text As String, ByVal Password As String) As String
    Dim passLen As Long
    Dim i As Long
    Dim passChr As Integer
    Dim passNdx As Long
    
    passLen = Len(Password)
    ' null passwords are invalid
    If passLen = 0 Then Err.Raise 5
    
    ' move password chars into an array of Integers to speed up code
    ReDim passChars(0 To passLen - 1) As Integer
    CopyMemory passChars(0), ByVal StrPtr(Password), passLen * 2
    
    ' this simple algorithm XORs each character of the string
    ' with a character of the password, but also modifies the
    ' password while it goes, to hide obvious patterns in the
    ' result string
    For i = 1 To Len(Text)
        ' get the next char in the password
        passChr = passChars(passNdx)
        ' encrypt one character in the string
        Mid$(Text, i, 1) = Chr$(Asc(Mid$(Text, i, 1)) Xor passChr)
        ' modify the character in the password (avoid overflow)
        passChars(passNdx) = (passChr + 17) And 255
        ' prepare to use next char in the password
        passNdx = (passNdx + 1) Mod passLen
    Next

    EncryptString = Text
End Function

Private Sub txtRegnKey_KeyDown(KeyCode As Integer, Shift As Integer)
'    If Shift = vbCtrlMask And KeyCode = 86 Then
'        txtRegnKey.Text = " Fix error in itemtrxfile table"
'    End If
    'm_textbox.SelStart = Len(m_textbox.Text)
End Sub

