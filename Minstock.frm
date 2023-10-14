VERSION 5.00
Begin VB.Form frmminstock 
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   2490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6270
   ControlBox      =   0   'False
   Icon            =   "Minstock.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2490
   ScaleWidth      =   6270
   Begin VB.TextBox txtitem 
      Alignment       =   2  'Center
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
      Left            =   165
      TabIndex        =   5
      Top             =   90
      Width           =   5430
   End
   Begin VB.ListBox LSTITEM 
      Height          =   1620
      Left            =   150
      TabIndex        =   7
      Top             =   495
      Width           =   5460
   End
   Begin VB.Frame Frame 
      Height          =   1530
      Left            =   135
      TabIndex        =   8
      Top             =   570
      Width           =   5505
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   495
         Left            =   75
         TabIndex        =   1
         Top             =   945
         Width           =   1230
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "<<PREVIOUS"
         Height          =   495
         Left            =   1470
         TabIndex        =   2
         Top             =   930
         Width           =   1170
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "NEXT >>"
         Height          =   495
         Left            =   2820
         TabIndex        =   3
         Top             =   930
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
         Height          =   360
         Left            =   1785
         MaxLength       =   4
         TabIndex        =   0
         Top             =   300
         Width           =   840
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Exit"
         Height          =   495
         Left            =   4200
         TabIndex        =   4
         Top             =   945
         Width           =   1200
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
         Left            =   135
         TabIndex        =   10
         Top             =   330
         Width           =   1470
      End
      Begin VB.Label LBLITEMCODE 
         BorderStyle     =   1  'Fixed Single
         Height          =   435
         Left            =   3720
         TabIndex        =   9
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Label LBLNO 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5685
      TabIndex        =   6
      Top             =   105
      Width           =   525
   End
End
Attribute VB_Name = "frmminstock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CMDEXIT_Click()
    MDIMAIN.PCTMENU.Enabled = True
    Unload Me
End Sub

Private Sub cmdnext_Click()
    Dim rstMINSTOCK As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrHand
    
    i = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From MINSTOCK ORDER BY SLNO, ITEM_NAME", db2, adOpenStatic, adLockReadOnly
    If Val(LBLNO.Caption) = RSTTRXFILE.RecordCount Then GoTo SKIP
    Do Until RSTTRXFILE.EOF
        i = i + 1
        If i = Val(LBLNO.Caption) + 1 Then
            Set rstMINSTOCK = New ADODB.Recordset
            rstMINSTOCK.Open "Select REORDER_QTY From ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                txtminqty.Text = rstMINSTOCK!REORDER_QTY
            End If
            TXTITEM.Text = RSTTRXFILE!ITEM_NAME
            LBLITEMCODE.Caption = RSTTRXFILE!ITEM_CODE
            rstMINSTOCK.Close
            Set rstMINSTOCK = Nothing
            LBLNO.Caption = Val(LBLNO.Caption) + 1
            GoTo SKIP
        End If
        RSTTRXFILE.MoveNext
    Loop
SKIP:
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    txtminqty.SetFocus

    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdOK_Click()
    Dim rstMINSTOCK As ADODB.Recordset
    On Error GoTo ErrHand
    
    If TXTITEM.Text = "" Then Exit Sub
    Set rstMINSTOCK = New ADODB.Recordset
    rstMINSTOCK.Open "Select REORDER_QTY, REMARKS From ITEMMAST WHERE ITEM_CODE = '" & LBLITEMCODE.Caption & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rstMINSTOCK.EOF And rstMINSTOCK.BOF) Then
        rstMINSTOCK!REORDER_QTY = Val(txtminqty.Text)
        'rstMINSTOCK!REMARKS = Val(txtminqty.Text)
        rstMINSTOCK.Update
    End If
    rstMINSTOCK.Close
    Set rstMINSTOCK = Nothing
    
    db2.Execute "Delete * from minstock WHERE ITEM_CODE = '" & LBLITEMCODE.Caption & "'"
    
    LBLNO.Caption = 0
    cmdnext.SetFocus
    cmdnext_Click
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdprev_Click()
    Dim rstMINSTOCK As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    On Error GoTo ErrHand
    
    i = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From MINSTOCK ORDER BY SLNO, ITEM_NAME", db2, adOpenStatic, adLockReadOnly
    If Val(LBLNO.Caption) = 1 Then GoTo SKIP
    Do Until RSTTRXFILE.EOF
        i = i + 1
        If i = Val(LBLNO.Caption) - 1 Then
            Set rstMINSTOCK = New ADODB.Recordset
            rstMINSTOCK.Open "Select REORDER_QTY From ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                txtminqty.Text = rstMINSTOCK!REORDER_QTY
            End If
            TXTITEM.Text = RSTTRXFILE!ITEM_NAME
            LBLITEMCODE.Caption = RSTTRXFILE!ITEM_CODE
            rstMINSTOCK.Close
            Set rstMINSTOCK = Nothing
            LBLNO.Caption = Val(LBLNO.Caption) - 1
            GoTo SKIP
        End If
        RSTTRXFILE.MoveNext
    Loop
SKIP:
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    txtminqty.SetFocus
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Command1_Click()
    Dim rstMINSTOCK As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTTRXMAST As ADODB.Recordset
    Dim i As Double
    Dim n As Integer
    Dim M As Integer
    On Error GoTo ErrHand
    
    i = 0
    n = 0
    db2.Execute "Delete * from minstock"
    'Set RSTTRXFILE = New ADODB.Recordset
    'RSTTRXFILE.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From RTRXFILE ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME , ITEMMAST.REMARKS FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE TRIM(ITEMMAST.REORDER_QTY) =  RTRXFILE.UNIT ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        n = n + 1
        Set rstMINSTOCK = New ADODB.Recordset
        rstMINSTOCK.Open "Select ITEM_CODE, ITEM_NAME From RTRXFILE WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstMINSTOCK.EOF And rstMINSTOCK.BOF) Then
            i = rstMINSTOCK.RecordCount
        End If
        rstMINSTOCK.Close
        Set rstMINSTOCK = Nothing
        
        Set rstMINSTOCK = New ADODB.Recordset
        rstMINSTOCK.Open "Select REORDER_QTY From ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstMINSTOCK.EOF And rstMINSTOCK.BOF) Then
            M = rstMINSTOCK!REORDER_QTY
        End If
        rstMINSTOCK.Close
        Set rstMINSTOCK = Nothing
        
        Set rstMINSTOCK = New ADODB.Recordset
        rstMINSTOCK.Open "Select * FROM MINSTOCK", db2, adOpenStatic, adLockOptimistic, adCmdText
        rstMINSTOCK.AddNew
        rstMINSTOCK!ITEM_CODE = RSTTRXFILE!ITEM_CODE
        rstMINSTOCK!ITEM_NAME = RSTTRXFILE!ITEM_NAME
        rstMINSTOCK!MINQTY = M
        rstMINSTOCK!SLNO = n
        rstMINSTOCK!Count = i
        rstMINSTOCK.Update

        rstMINSTOCK.Close
        Set rstMINSTOCK = Nothing
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
    Dim rstMINSTOCK As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTTRXMAST As ADODB.Recordset
    Dim i As Double
    Dim n As Integer
    Dim M As Integer
    On Error GoTo ErrHand
    
    i = 0
    n = 0

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From MINSTOCK ORDER BY COUNT", db2, adOpenStatic, adLockOptimistic, adCmdText
    i = RSTTRXFILE.RecordCount
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!SLNO = i
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
        i = i - 1
    Loop
    rstMINSTOCK.Close
    Set rstMINSTOCK = Nothing
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim rstMINSTOCK As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From MINSTOCK ORDER BY SLNO, ITEM_NAME", db2, adOpenStatic, adLockReadOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        Set rstMINSTOCK = New ADODB.Recordset
        rstMINSTOCK.Open "Select REORDER_QTY From ITEMMAST WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly
        If Not (rstMINSTOCK.EOF And rstMINSTOCK.BOF) Then
            txtminqty.Text = rstMINSTOCK!REORDER_QTY
        End If
        TXTITEM.Text = RSTTRXFILE!ITEM_NAME
        LBLITEMCODE.Caption = RSTTRXFILE!ITEM_CODE
        rstMINSTOCK.Close
        Set rstMINSTOCK = Nothing
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LSTITEM.Visible = False
    Frame.Enabled = True
    LBLNO.Caption = 1
    Me.Width = 7000
    Me.Height = 3000
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub LSTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select ITEM_CODE, ITEM_NAME, REORDER_QTY From ITEMMAST WHERE ITEM_NAME = '" & LSTITEM.Text & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                txtminqty.Text = RSTTRXFILE!REORDER_QTY
                TXTITEM.Text = RSTTRXFILE!ITEM_NAME
                LBLITEMCODE.Caption = RSTTRXFILE!ITEM_CODE
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                LBLNO.Caption = 0
            Else
                TXTITEM.Text = ""
                LBLITEMCODE.Caption = ""
                LBLNO.Caption = 0
            End If
            LSTITEM.Visible = False
            Frame.Enabled = True
            txtminqty.SetFocus
    End Select
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As String
    Dim RSTTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then Exit Sub
            'txtminqty.SetFocus
        'Case vbKeyF3
            LSTITEM.Clear
            i = TXTITEM.Text 'InputBox("Enter the Search text", "SEARCH")
            If i = "" Then Exit Sub
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST WHERE ITEM_NAME Like '" & i & "%'", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                If RSTTRXFILE.RecordCount = 1 Then Exit Do
                LSTITEM.AddItem RSTTRXFILE!ITEM_NAME
                RSTTRXFILE.MoveNext
            Loop
            If RSTTRXFILE.RecordCount > 1 Then
                TXTITEM.Text = ""
                txtminqty.Text = ""
                LBLITEMCODE.Caption = ""
                LSTITEM.Visible = True
                Frame.Enabled = False
                LSTITEM.SetFocus
            End If
            If RSTTRXFILE.RecordCount = 1 Then
                TXTITEM.Text = RSTTRXFILE!ITEM_NAME
                txtminqty.SetFocus
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
     End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub txtminqty_GotFocus()
    txtminqty.SelStart = 0
    txtminqty.SelLength = Len(txtminqty.Text)
End Sub

Private Sub txtminqty_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK.SetFocus
    End Select
End Sub

Private Sub txtminqty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

