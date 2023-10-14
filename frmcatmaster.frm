VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCatmaster 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Details"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   ClipControls    =   0   'False
   Icon            =   "frmcatmaster.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11685
   Begin VB.CommandButton CmdDisplay 
      Caption         =   "&Display"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   7680
      TabIndex        =   1
      Top             =   60
      Width           =   1335
   End
   Begin VB.TextBox tXTMEDICINE 
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
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   255
      Width           =   2760
   End
   Begin VB.CommandButton CmdExit 
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
      Height          =   540
      Left            =   10575
      TabIndex        =   3
      Top             =   60
      Width           =   1065
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete category (Ctrl +D)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   9060
      TabIndex        =   2
      Top             =   60
      Width           =   1485
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Category (Ctrl +I)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   5970
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   7635
      Left            =   45
      TabIndex        =   6
      Top             =   570
      Width           =   11610
      Begin VB.TextBox TXTsample 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   210
         TabIndex        =   7
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   7470
         Left            =   30
         TabIndex        =   4
         Top             =   120
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   13176
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColor       =   15985374
         BackColorFixed  =   0
         ForeColorFixed  =   8438015
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Category Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   9
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2760
   End
End
Attribute VB_Name = "FrmCatmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdDelete_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    Dim i As Integer
    
    If GRDSTOCK.rows <= 1 Then Exit Sub
    On Error GoTo ErrHand
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from RTRXFILE where CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' ", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & " Since Transactions is Available", vbCritical, "DELETE"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' ", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & " Since Transactions is Available", vbCritical, "DELETE"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & " ****", vbYesNo + vbDefaultButton2, "DELETE....") = vbNo Then
        GRDSTOCK.SetFocus
        Exit Sub
    End If
    
    db.Execute ("DELETE from CATEGORY where CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    
    Dim selrow As Integer
    selrow = GRDSTOCK.Row
    For i = selrow To GRDSTOCK.rows - 2
        GRDSTOCK.TextMatrix(selrow, 0) = i
        GRDSTOCK.TextMatrix(selrow, 1) = GRDSTOCK.TextMatrix(i + 1, 1)
        GRDSTOCK.TextMatrix(selrow, 2) = GRDSTOCK.TextMatrix(i + 1, 2)
        GRDSTOCK.TextMatrix(selrow, 3) = GRDSTOCK.TextMatrix(i + 1, 3)
        selrow = selrow + 1
    Next i
    GRDSTOCK.rows = GRDSTOCK.rows - 1
    GRDSTOCK.SetFocus
    Exit Sub
   
ErrHand:
    MsgBox err.Description
End Sub


Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdnew_Click()
    On Error GoTo ErrHand
    If GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 2) <> "" Then
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 0) = GRDSTOCK.rows - 1
        
'        Dim TRXMAST As ADODB.Recordset
'        On Error GoTo ErrHand
'
'        Set TRXMAST = New ADODB.Recordset
'        TRXMAST.Open "Select MAX(CONVERT(Veh_Code, SIGNED INTEGER)) From Veh_Master ", db, adOpenStatic, adLockReadOnly
'        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'            If IsNull(TRXMAST.Fields(0)) Then
'                GRDSTOCK.TextMatrix(GRDSTOCK.Rows - 1, 1) = 1
'            Else
'                GRDSTOCK.TextMatrix(GRDSTOCK.Rows - 1, 1) = Val(TRXMAST.Fields(0)) + 1
'            End If
'        End If
'        TRXMAST.Close
'        Set TRXMAST = Nothing
    End If
    TXTsample.Visible = False
    'GRDSTOCK.TopRow = GRDSTOCK.Rows - 1
    GRDSTOCK.Row = GRDSTOCK.rows - 1
    GRDSTOCK.Col = 2
    GRDSTOCK.SetFocus
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case 73
                Call cmdnew_Click
            Case 68
                Call CmdDelete_Click
        End Select
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHand
    
    
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "CATEGORY"
    GRDSTOCK.TextMatrix(0, 2) = "COOLIE"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 5000
    GRDSTOCK.ColWidth(2) = 1800
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 4
    
    Call Fillgrid
    Me.Left = 500
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDSTOCK.Col
                    Case 1
                        TXTsample.MaxLength = 50
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 2
                        If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)) = "" Then Exit Sub
                        TXTsample.MaxLength = 5
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                End Select
            End If
        
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    
    TXTsample.Visible = False
    
End Sub


Private Function Fillgrid()
    
    Dim rststock As ADODB.Recordset
    Dim rstopstock As ADODB.Recordset
    Dim i As Long
    
    
    On Error GoTo ErrHand
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    db.Execute "Update category set category ='' where isnull(category)"
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM category WHERE category Like '%" & tXTMEDICINE.text & "%' ORDER BY category", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        'GRDSTOCK.FixedCols = 3
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!Category
        GRDSTOCK.TextMatrix(i, 2) = IIf(IsNull(rststock!COOLIE), 0, rststock!COOLIE)
        
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    'Call Toatal_value
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub tXTMEDICINE_Change()
    Call Fillgrid
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            'TxtCode.SetFocus
            Call CmDDisplay_Click
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 1 'CATEGORY
                    If Trim(TXTsample.text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from CATEGORY where CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rststock.EOF And rststock.BOF) Then
                        rststock.AddNew
                        rststock!Category = Trim(TXTsample.text)
                    Else
                        rststock!Category = Trim(TXTsample.text)
                    End If
                    rststock.Update
                    rststock.Close
                    Set rststock = Nothing
                    
'                    db.Execute "UPDATE TRXFILE set CATEGORY = '" & Trim(TXTsample.text) & "' WHERE CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
'                    db.Execute "UPDATE RTRXFILE set CATEGORY = '" & Trim(TXTsample.text) & "' WHERE CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' "
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 2  'COOLIE
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from CATEGORY where CATEGORY = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!COOLIE = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 1
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        Case 2
             Select Case KeyAscii
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub


