VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmitemmasterwo 
   Caption         =   "Item Creation"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   ControlBox      =   0   'False
   Icon            =   "frmitemmasterwo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   8265
   Begin VB.TextBox TxtProduct 
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
      Left            =   1320
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.TextBox TxtItemcode 
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
      Left            =   1320
      MaxLength       =   10
      TabIndex        =   19
      Top             =   420
      Width           =   1455
   End
   Begin VB.Frame FRAME 
      Height          =   3285
      Left            =   30
      TabIndex        =   1
      Top             =   1020
      Width           =   8205
      Begin VB.TextBox TXTITEM 
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
         Left            =   1725
         TabIndex        =   14
         Top             =   195
         Width           =   6180
      End
      Begin VB.Frame FrmeCompany 
         BorderStyle     =   0  'None
         Height          =   1710
         Left            =   4020
         TabIndex        =   9
         Top             =   435
         Width           =   4020
         Begin VB.TextBox txtcompany 
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
            ForeColor       =   &H00004080&
            Height          =   330
            Left            =   1035
            TabIndex        =   11
            Top             =   165
            Width           =   2895
         End
         Begin VB.CheckBox chknewcomp 
            Caption         =   "&New Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   1035
            TabIndex        =   10
            Top             =   1260
            Width           =   1695
         End
         Begin MSDataListLib.DataList Datacompany 
            Height          =   645
            Left            =   1035
            TabIndex        =   12
            Top             =   510
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   1138
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16512
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
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
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
            Left            =   0
            TabIndex        =   13
            Top             =   210
            Width           =   960
         End
      End
      Begin VB.CommandButton CMDDELETE 
         BackColor       =   &H00400000&
         Caption         =   "&Delete"
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
         Left            =   7155
         MaskColor       =   &H80000007&
         TabIndex        =   8
         Top             =   2670
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin VB.CheckBox chknewcategory 
         Caption         =   "N&ew Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1050
         TabIndex        =   7
         Top             =   1695
         Width           =   1725
      End
      Begin VB.TextBox txtcategory 
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
         ForeColor       =   &H00004080&
         Height          =   330
         Left            =   1050
         TabIndex        =   6
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TxtMinQty 
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
         Left            =   1260
         TabIndex        =   5
         Top             =   2175
         Width           =   1020
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
         Left            =   120
         MaskColor       =   &H80000007&
         TabIndex        =   4
         Top             =   2685
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
         Left            =   1140
         MaskColor       =   &H80000007&
         TabIndex        =   3
         Top             =   2685
         UseMaskColor    =   -1  'True
         Width           =   975
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
         Left            =   2220
         MaskColor       =   &H80000007&
         TabIndex        =   2
         Top             =   2670
         UseMaskColor    =   -1  'True
         Width           =   930
      End
      Begin MSDataListLib.DataList Datacategory 
         Height          =   645
         Left            =   1050
         TabIndex        =   15
         Top             =   945
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1138
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16512
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
         Caption         =   "Item Description"
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
         Left            =   75
         TabIndex        =   18
         Top             =   195
         Width           =   1800
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
         Index           =   2
         Left            =   45
         TabIndex        =   17
         Top             =   645
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MIN QTY"
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
         TabIndex        =   16
         Top             =   2220
         Width           =   1050
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1320
      TabIndex        =   20
      Top             =   780
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
      Caption         =   "ITEM CODE"
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
      Left            =   195
      TabIndex        =   22
      Top             =   420
      Width           =   1050
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
      TabIndex        =   21
      Top             =   0
      Width           =   6300
   End
End
Attribute VB_Name = "frmitemmasterwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean
Dim COMPANYFLAG As Boolean
Dim CATEGORYFLAG As Boolean
Dim CLOSEALL As Integer
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset
Dim RSTCATEGORY As New ADODB.Recordset
'Dim rstTMP As New ADODB.Recordset
'Dim TMPFLAG As Boolean 'TMP

Private Sub chknewcategory_Click()
    On Error Resume Next
    txtcategory.SetFocus
End Sub

Private Sub chknewcomp_Click()
    On Error Resume Next
    txtcompany.SetFocus
End Sub

Private Sub cmdcancel_Click()

    TxtProduct.Text = ""
    TXTITEM.Text = ""
    txtcategory.Text = ""
    txtcompany.Text = ""
    TxtMinQty.Text = ""
    Set DataList2.RowSource = Nothing
    TxtItemcode.Enabled = True
    DataList2.Enabled = True
    FRAME.Visible = False
    TxtProduct.Visible = False
    DataList2.Visible = False
    TxtItemcode.SetFocus
    chknewcategory.Value = 0
    chknewcomp.Value = 0
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If TXTITEM.Text = "" Then
        MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    
     If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And Trim(txtcompany.Text) = "" Then
        MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
        txtcompany.SetFocus
        Exit Sub
    End If
    If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
        MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
        txtcompany.SetFocus
        Exit Sub
    End If
    
    If Trim(txtcategory.Text) = "" Then
        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
        txtcategory.SetFocus
        Exit Sub
    End If
    
    If chknewcategory.Value = 0 And Datacategory.BoundText = "" Then
        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
        txtcategory.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRHAND
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMASTWO WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.Value = 1 Then RSTITEMMAST!CATEGORY = txtcategory.Text Else RSTITEMMAST!CATEGORY = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.Value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Datacompany.BoundText
        RSTITEMMAST!REMARKS = ""
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        RSTITEMMAST!BIN_LOCATION = ""
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!ITEM_CODE = Trim(TxtItemcode.Text)
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.Value = 1 Then RSTITEMMAST!CATEGORY = txtcategory.Text Else RSTITEMMAST!CATEGORY = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.Value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Trim(Datacompany.BoundText)
        RSTITEMMAST!REMARKS = ""
        RSTITEMMAST!REORDER_QTY = Val(TxtMinQty.Text)
        RSTITEMMAST!BIN_LOCATION = ""
        RSTITEMMAST!ITEM_COST = 0
        RSTITEMMAST!MRP = 0
        RSTITEMMAST!SALES_TAX = 0
        RSTITEMMAST!PTR = 0
        RSTITEMMAST!CST = 0
        RSTITEMMAST!OPEN_QTY = 0
        RSTITEMMAST!OPEN_VAL = 0
        RSTITEMMAST!RCPT_QTY = 0
        RSTITEMMAST!RCPT_VAL = 0
        RSTITEMMAST!ISSUE_QTY = 0
        RSTITEMMAST!ISSUE_VAL = 0
        RSTITEMMAST!CLOSE_QTY = 0
        RSTITEMMAST!CLOSE_VAL = 0
        RSTITEMMAST!DAM_QTY = 0
        RSTITEMMAST!DAM_VAL = 0
        RSTITEMMAST!DISC = 0
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT [MANUFACTURER] FROM MANUFACT WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!MANUFACTURER = Trim(txtcompany.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT [CATEGORY] FROM CATEGORY WHERE CATEGORY = '" & Trim(txtcategory.Text) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!CATEGORY = Trim(txtcategory.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * from RTRXFILEWO WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        RSTITEMMAST!MFGR = Trim(txtcompany.Text)
        RSTITEMMAST!CATEGORY = Trim(txtcategory.Text)
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    cmdcancel_Click
Exit Sub
eRRHAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtMinQty.SetFocus
    End Select
End Sub

Private Sub Datacategory_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo eRRHAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT [CATEGORY] FROM ITEMMASTWO WHERE ITEM_CODE = '" & Datacategory.BoundText & "'", DB2, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        TXTCATEGORY.Text = RSTITEMMAST!CATEGORY
'    End If
    txtcategory.Text = Datacategory.BoundText
    Datacategory.Text = txtcategory.Text
        
    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
        FrmeCompany.Visible = False
    Else
        FrmeCompany.Visible = True
    End If
            
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)

    
End Sub

Private Sub Datacategory_GotFocus()
    Datacategory.Text = txtcategory.Text
End Sub

Private Sub Datacompany_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo eRRHAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT [MANUFACTURER] FROM ITEMMASTWO WHERE ITEM_CODE = '" & Datacompany.BoundText & "'", DB2, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        txtcompany.Text = RSTITEMMAST!MANUFACTURER
'    End If
    txtcompany.Text = Datacompany.BoundText
    Datacompany.Text = txtcompany.Text
    
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub Datacompany_GotFocus()
    Datacompany.Text = txtcompany.Text
    'Call Datacompany_Click
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRHAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT [ITEM_CODE] FROM ITEMMASTWO WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TxtItemcode.Text = RSTITEMMAST!ITEM_CODE
            End If
            TxtProduct.Visible = False
            DataList2.Visible = False
            TxtItemcode.SetFocus
        Case vbKeyEscape
            TxtProduct.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    Call TXTCOMPANY_Change
    TxtItemcode.SetFocus
End Sub

Private Sub Form_Load()

    On Error GoTo eRRHAND
    
    REPFLAG = True
    COMPANYFLAG = True
    CATEGORYFLAG = True
    Call txtcategory_Change
    'TMPFLAG = True
    CLOSEALL = 1
    Width = 8385
    Height = 4575
    Left = 3500
    Top = 1900
    FRAME.Visible = False
    'txtunit.Visible = False
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If REPFLAG = False Then RSTREP.Close
        If COMPANYFLAG = False Then RSTCOMPANY.Close
        If CATEGORYFLAG = False Then RSTCATEGORY.Close
        'If TMPFLAG = False Then rstTMP.Close
        MDIMAIN.Enabled = True
        'FrmCrimedata.Enabled = True
    End If
   Cancel = CLOSEALL
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then
                MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            txtcategory.SetFocus
    End Select
    
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtItemcode_GotFocus()
    TxtItemcode.SelStart = 0
    TxtItemcode.SelLength = Len(TxtItemcode.Text)
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRHAND
            If Trim(TxtItemcode.Text) = "" Then Exit Sub
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ITEMMASTWO WHERE ITEM_CODE = '" & Trim(TxtItemcode.Text) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTITEM.Text = RSTITEMMAST!ITEM_NAME
                txtcategory.Text = RSTITEMMAST!CATEGORY
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!MANUFACTURER), "", RSTITEMMAST!MANUFACTURER)
                TxtMinQty.Text = IIf(IsNull(RSTITEMMAST!REORDER_QTY), 0, RSTITEMMAST!REORDER_QTY)
                Datacategory.Text = txtcategory.Text
                Call Datacategory_Click
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            TxtItemcode.Enabled = False
            FRAME.Visible = True
            TXTITEM.SetFocus
        Case 114
            TxtProduct.Visible = True
            DataList2.Visible = True
            TxtProduct.SetFocus
        Case vbKeyEscape
            Call CMDEXIT_Click
    End Select
Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtminqty_GotFocus()
    If TxtMinQty.Text = "" Then
        TxtMinQty.Text = 1
    End If
    TxtMinQty.SelStart = 0
    TxtMinQty.SelLength = Len(TxtMinQty.Text)
End Sub

Private Sub TxtMinQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
        Case vbKeyEscape
            If FrmeCompany.Visible = True Then
                txtcompany.SetFocus
            Else
                txtcategory.SetFocus
            End If
    End Select
End Sub

Private Sub TxtMinQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtProduct_Change()
    On Error GoTo eRRHAND
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & TxtProduct.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] FROM ITEMMASTWO  WHERE ITEM_NAME Like '" & TxtProduct.Text & "%'ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TxtProduct.SelStart = 0
    TxtProduct.SelLength = Len(TxtProduct.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TxtProduct.Text = "" Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            TxtProduct.Visible = False
            DataList2.Visible = False
            TxtItemcode.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTCOMPANY_Change()
    On Error GoTo eRRHAND
    
    Set Datacompany.RowSource = Nothing
    If COMPANYFLAG = True Then
        RSTCOMPANY.Open "Select DISTINCT [MANUFACTURER] From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY [MANUFACTURER]", db2, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    Else
        RSTCOMPANY.Close
        RSTCOMPANY.Open "Select DISTINCT [MANUFACTURER] From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY [MANUFACTURER]", db2, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    End If
    Set Datacompany.RowSource = RSTCOMPANY
    Datacompany.ListField = "MANUFACTURER"
    Datacompany.BoundColumn = "MANUFACTURER"
    
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTCOMPANY_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.Text)
End Sub

Private Sub TXTCOMPANY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            '''''''If txtcompany.Text = "" Then Exit Sub
            Datacompany.SetFocus
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTCOMPANY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRHAND
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "SELECT [MANUFACTURER] FROM ITEMMASTWO WHERE ITEM_CODE ='" & Trim(TXTITEMCODE.Text) & "'", DB2, adOpenStatic, adLockReadOnly
'            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                txtcompany.Text = RSTITEMMAST!MANUFACTURER
'            Else
'            End If
            If txtcompany.Text = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If
            If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If
            
            TxtMinQty.SetFocus
        Case vbKeyEscape
            txtcompany.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub txtcategory_Change()
    On Error GoTo eRRHAND
    If CATEGORYFLAG = True Then
        RSTCATEGORY.Open "Select DISTINCT [CATEGORY] From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY [CATEGORY]", db2, adOpenStatic, adLockReadOnly
        ''RSTCATEGORY.Open "Select DISTINCT [CATEGORY] From CATEGORY ORDER BY [CATEGORY]", DB2, adOpenStatic, adLockReadOnly
        CATEGORYFLAG = False
    Else
        RSTCATEGORY.Close
        RSTCATEGORY.Open "Select DISTINCT [CATEGORY] From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY [CATEGORY]", db2, adOpenStatic, adLockReadOnly
        CATEGORYFLAG = False
    End If
    Set Datacategory.RowSource = RSTCATEGORY
    Datacategory.ListField = "CATEGORY"
    Datacategory.BoundColumn = "CATEGORY"
    
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            '''''If txtcategory.Text = "" Then Exit Sub
            Datacategory.SetFocus
        Case vbKeyEscape
            TXTITEM.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRHAND
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "SELECT [CATEGORY] FROM ITEMMASTWO WHERE ITEM_CODE ='" & Trim(TXTITEMCODE.Text) & "'", DB2, adOpenStatic, adLockReadOnly
'            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'                txtcategory.Text = RSTITEMMAST!CATEGORY
'            Else
'            End If
            If txtcategory.Text = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If
            
            If chknewcategory.Value = 0 And Datacategory.BoundText = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If
            If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
                FrmeCompany.Visible = False
                TxtMinQty.SetFocus
            Else
                FrmeCompany.Visible = True
                txtcompany.SetFocus
            End If
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox (Err.Description)
End Sub

Private Sub CmdDelete_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    If TxtItemcode.Text = "" Then Exit Sub
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from RTRXFILEWO where RTRXFILEWO.ITEM_CODE = '" & TxtItemcode.Text & "'", db2, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    If i <> 0 Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Stock is Available", vbCritical, "DELETING ITEM...."
        Exit Sub
    End If
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & TXTITEM.Text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    db2.Execute ("DELETE from RTRXFILEWO where RTRXFILEWO.ITEM_CODE = '" & TxtItemcode.Text & "'")
    db2.Execute ("DELETE from [PRODLINK] where PRODLINK.ITEM_CODE = '" & TxtItemcode.Text & "'")
    db2.Execute ("DELETE FROM ITEMMASTWO where ITEMMASTWO.ITEM_CODE = '" & TxtItemcode.Text & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
    MsgBox "ITEM " & TXTITEM.Text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    Call cmdcancel_Click
    Exit Sub
   
eRRHAND:
    MsgBox Err.Description
End Sub


