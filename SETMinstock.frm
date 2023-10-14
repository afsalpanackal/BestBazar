VERSION 5.00
Begin VB.Form frmminstock 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   1710
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "SETMinstock.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1710
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   1695
      Left            =   45
      TabIndex        =   3
      Top             =   -30
      Width           =   4230
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   825
         TabIndex        =   1
         Top             =   1050
         Width           =   1200
      End
      Begin VB.TextBox txtminqty 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   2415
         MaxLength       =   4
         TabIndex        =   0
         Top             =   690
         Width           =   840
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   405
         Left            =   2115
         TabIndex        =   2
         Top             =   1050
         Width           =   1215
      End
      Begin VB.Label LBLITEMNAME 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   135
         TabIndex        =   5
         Top             =   240
         Width           =   4005
      End
      Begin VB.Label Label1 
         Caption         =   "Minimum Stock"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   870
         TabIndex        =   4
         Top             =   750
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmminstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    frmstockless.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim RSTMINSTOCK As ADODB.Recordset
    Dim i As Long
    Dim SN As Integer
    On Error GoTo eRRhAND
    
    Set RSTMINSTOCK = New ADODB.Recordset
    RSTMINSTOCK.Open "SELECT REORDER_QTY FROM ITEMMAST WHERE ITEM_CODE = '" & frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTMINSTOCK.EOF And RSTMINSTOCK.BOF) Then
        RSTMINSTOCK!REORDER_QTY = Val(TxtMinQty.Text)
        frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 5) = Val(TxtMinQty.Text)
        If Val(frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 5)) <= Val(frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 3)) Then
            SN = frmstockless.grdSTOCKLESS.Row
            For i = SN To frmstockless.grdSTOCKLESS.Rows - 2
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 0) = frmstockless.grdSTOCKLESS.TextMatrix(i, 0)
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 1) = frmstockless.grdSTOCKLESS.TextMatrix(i + 1, 1)
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 2) = frmstockless.grdSTOCKLESS.TextMatrix(i + 1, 2)
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 3) = frmstockless.grdSTOCKLESS.TextMatrix(i + 1, 3)
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 4) = frmstockless.grdSTOCKLESS.TextMatrix(i + 1, 4)
                frmstockless.grdSTOCKLESS.TextMatrix(SN, 5) = frmstockless.grdSTOCKLESS.TextMatrix(i + 1, 5)
                SN = SN + 1
            Next i
            frmstockless.grdSTOCKLESS.Rows = frmstockless.grdSTOCKLESS.Rows - 1
        End If
        RSTMINSTOCK.Update
    End If
    RSTMINSTOCK.Close
    Set RSTMINSTOCK = Nothing
    Call cmdexit_Click
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo eRRhAND
    
    lblitemname.Caption = frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 2)
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT REORDER_QTY FROM ITEMMAST WHERE ITEM_CODE = '" & frmstockless.grdSTOCKLESS.TextMatrix(frmstockless.grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        TxtMinQty.Text = RSTTRXFILE!REORDER_QTY
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Me.Width = 4500
    Me.Height = 1700
    Me.Left = 5500
    Me.Top = 2000
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub


Private Sub txtminqty_GotFocus()
    TxtMinQty.SelStart = 0
    TxtMinQty.SelLength = Len(TxtMinQty.Text)
End Sub

Private Sub TxtMinQty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK.SetFocus
    End Select
End Sub

Private Sub TxtMinQty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

