VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmIncMast 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Income Head Creation"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   ControlBox      =   0   'False
   Icon            =   "frmIncMast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   7155
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
      Left            =   1815
      MaxLength       =   34
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00E0E0E0&
      Height          =   1815
      Left            =   90
      TabIndex        =   8
      Top             =   1245
      Width           =   6765
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
         Left            =   255
         MaskColor       =   &H80000007&
         TabIndex        =   11
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   1260
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
         Left            =   1695
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
         Left            =   3645
         MaskColor       =   &H80000007&
         TabIndex        =   4
         Top             =   945
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
         Left            =   4680
         MaskColor       =   &H80000007&
         TabIndex        =   5
         Top             =   945
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
         Left            =   5715
         MaskColor       =   &H80000007&
         TabIndex        =   6
         Top             =   945
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Income Head"
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
         TabIndex        =   9
         Top             =   255
         Width           =   1515
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
      Left            =   1815
      MaxLength       =   3
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1815
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
      TabIndex        =   10
      Top             =   45
      Width           =   6300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Income Code"
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
      TabIndex        =   7
      Top             =   465
      Width           =   1560
   End
End
Attribute VB_Name = "frmIncMast"
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
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXINCOME WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "DELETE!!!!") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME FOR INCOME HEAD", vbOKOnly, "INCOME MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRhAND
    Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
    Txtsuplcode.Tag = "741" & Txtsuplcode.Text
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!act_code = Txtsuplcode.Tag
    End If
    RSTITEMMAST!ACT_NAME = Trim(txtsupplier.Text)
    RSTITEMMAST!OPEN_DB = 0
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
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    Frame.Visible = False
    txtsupplier.Text = ""
    Txtsuplcode.Enabled = True
Exit Sub
eRRhAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRhAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ACT_CODE FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.Text = Mid(RSTITEMMAST!act_code, Len(RSTITEMMAST!act_code) - 2, 3)
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
    If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    Me.Width = 7000
    Me.Height = 3600
    Me.Left = 2500
    Me.Top = 1900
    Frame.Visible = False
    'txtunit.Visible = False
    On Error GoTo eRRhAND
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(ACT_CODE) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='741')And (LENGTH(ACT_CODE)>3)", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "001", Mid(TRXMAST.Fields(0), 4) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
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

Private Sub txtsupplier_GotFocus()
    txtsupplier.SelStart = 0
    txtsupplier.SelLength = Len(txtsupplier.Text)
   
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtsupplier.Text = "" Then
                MsgBox "ENTER NAME FOR INCOME HEAD", vbOKOnly, "INCOME MASTER"
                txtsupplier.SetFocus
                Exit Sub
            End If
         CmdSave.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
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
            'If Val(Txtsuplcode.Text) = 1 Then Exit Sub
            On Error GoTo eRRhAND
            Txtsuplcode.Text = Format(Txtsuplcode.Text, "000")
            Txtsuplcode.Tag = "741" & Txtsuplcode.Text
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.Text = RSTITEMMAST!ACT_NAME
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
            Call CMDEXIT_Click
    End Select
Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Txtsuplcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplist_Change()
    On Error GoTo eRRhAND
    If REPFLAG = True Then
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='741')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='741')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


