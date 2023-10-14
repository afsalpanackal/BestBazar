VERSION 5.00
Begin VB.Form frmOPCash 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Opening Cash"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6990
   ClipControls    =   0   'False
   Icon            =   "frmOPCash.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   6990
   Begin VB.Frame FRAME 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   135
      TabIndex        =   3
      Top             =   60
      Width           =   6765
      Begin VB.TextBox txtsupplier 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   525
         Left            =   1695
         MaxLength       =   34
         TabIndex        =   0
         Top             =   225
         Width           =   2040
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
         Height          =   555
         Left            =   4080
         MaskColor       =   &H80000007&
         TabIndex        =   1
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1290
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
         Height          =   555
         Left            =   5415
         MaskColor       =   &H80000007&
         TabIndex        =   2
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Cash"
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
         TabIndex        =   4
         Top             =   375
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmOPCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "Please enter the Opening Cash", vbOKOnly, "OP Cash Entry"
        txtsupplier.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Errhand
    db.BeginTrans
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '111001'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!ACT_CODE = "111001"
    End If
    RSTITEMMAST!ACT_NAME = "Cash on Hand"
    RSTITEMMAST!OPEN_DB = Val(txtsupplier.Text)
    RSTITEMMAST!OPEN_CR = 0
    RSTITEMMAST!YTD_DB = 0
    RSTITEMMAST!YTD_CR = 0
    RSTITEMMAST!CREATE_DATE = Date
    RSTITEMMAST!C_USER_ID = "SM"
    RSTITEMMAST!MODIFY_DATE = Date
    RSTITEMMAST!M_USER_ID = "SM"
    
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '222222'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!ACT_CODE = "222222"
    End If
    RSTITEMMAST!ACT_NAME = "Cash Capital"
    RSTITEMMAST!OPEN_DB = 0
    RSTITEMMAST!OPEN_CR = Val(txtsupplier.Text)
    RSTITEMMAST!YTD_DB = 0
    RSTITEMMAST!YTD_CR = 0
    RSTITEMMAST!CREATE_DATE = Date
    RSTITEMMAST!C_USER_ID = "SM"
    RSTITEMMAST!MODIFY_DATE = Date
    RSTITEMMAST!M_USER_ID = "SM"
    
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    db.CommitTrans
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "OP Cash Entry"
Exit Sub
Errhand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub Form_Activate()
    'If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim RSTITEMMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    Me.Width = 7000
    Me.Height = 3600
    Me.Left = 2500
    Me.Top = 1900
    'txtunit.Visible = False
    On Error GoTo Errhand
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '111001'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        txtsupplier.Text = IIf(IsNull(RSTITEMMAST!OPEN_DB), 0, RSTITEMMAST!OPEN_DB)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Sub
Errhand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub txtsupplier_GotFocus()
    txtsupplier.SelStart = 0
    txtsupplier.SelLength = Len(txtsupplier.Text)
   
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtsupplier.Text = "" Then
                MsgBox "ENTER NAME FOR EXPENSE HEAD", vbOKOnly, "EXPENSE MASTER"
                txtsupplier.SetFocus
                Exit Sub
            End If
         CmdSave.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("."), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
