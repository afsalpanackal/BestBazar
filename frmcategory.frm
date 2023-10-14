VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCatmast 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Category Master"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   ControlBox      =   0   'False
   Icon            =   "frmcategory.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   6585
   Begin VB.Frame FRAME 
      BackColor       =   &H00E0E0E0&
      Height          =   3105
      Left            =   90
      TabIndex        =   7
      Top             =   240
      Width           =   6450
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
         Left            =   1410
         MaxLength       =   34
         TabIndex        =   0
         Top             =   120
         Width           =   4950
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00400000&
         Caption         =   "&DELETE"
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
         TabIndex        =   3
         Top             =   2565
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
         Left            =   1425
         MaxLength       =   6
         TabIndex        =   2
         Top             =   2130
         Width           =   1830
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
         Left            =   3405
         MaskColor       =   &H80000007&
         TabIndex        =   4
         Top             =   2565
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
         Left            =   4440
         MaskColor       =   &H80000007&
         TabIndex        =   5
         Top             =   2565
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
         Left            =   5475
         MaskColor       =   &H80000007&
         TabIndex        =   6
         Top             =   2565
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1620
         Left            =   1410
         TabIndex        =   1
         Top             =   480
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
         Caption         =   "Category"
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
         Left            =   285
         TabIndex        =   10
         Top             =   120
         Width           =   1560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Coolie"
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
         Left            =   360
         TabIndex        =   8
         Top             =   2160
         Width           =   1515
      End
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Esc to exit"
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
      TabIndex        =   9
      Top             =   45
      Width           =   6300
   End
End
Attribute VB_Name = "frmCatmast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo Errhand
    
    If DataList2.BoundText = "" Then Exit Sub
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT CATEGORY From ITEMMAST WHERE CATEGORY = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "DELETE!!!!") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM CATEGORY WHERE CATEGORY = '" & DataList2.BoundText & "'")
    txtsupplist.Text = ""
    txtsupplier.Text = ""
    'Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
Errhand:
    MsgBox (Err.Description)
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If DataList2.BoundText = "" Then
        MsgBox "PLEASE SELECT CATEGORY", vbOKOnly, "CATEGORY MASTER"
        txtsupplist.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Errhand
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CATEGORY WHERE CATEGORY = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!Category = DataList2.BoundText
    End If
    RSTITEMMAST!COOLIE = Val(txtsupplier.Text)
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "CATEGORY CREATION"
'    txtsupplist.Text = ""
'    txtsupplier.Text = ""
Exit Sub
Errhand:
    MsgBox (Err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtsupplier.SetFocus
    End Select
End Sub

Private Sub DataList2_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If DataList2.BoundText = "" Then
        MsgBox "PLEASE SELECT CATEGORY", vbOKOnly, "CATEGORY MASTER"
        txtsupplist.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Errhand
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CATEGORY WHERE CATEGORY = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        txtsupplier.Text = RSTITEMMAST!COOLIE
    Else
        txtsupplier.Text = ""
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
Exit Sub
Errhand:
    MsgBox (Err.Description)
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
             txtsupplier.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    'Me.Width = 7000
    'Me.Height = 3600
    Me.Left = 2500
    Me.Top = 1900
    'FRAME.Visible = False
    'txtunit.Visible = False
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
             CmdSave.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub txtsupplist_Change()
    On Error GoTo Errhand
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & Me.txtsupplist.Text & "%'ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & Me.txtsupplist.Text & "%'ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
        'RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic,adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "CATEGORY"
    DataList2.BoundColumn = "CATEGORY"
    
    Exit Sub
Errhand:
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
            Unload Me
            
    End Select
    Exit Sub
Errhand:
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


