VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmVehMaster 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vehicle Details"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11685
   ClipControls    =   0   'False
   Icon            =   "FrmVeh_Master.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   11685
   Begin VB.ComboBox CmbUnit2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "FrmVeh_Master.frx":030A
      Left            =   8880
      List            =   "FrmVeh_Master.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   105
      Width           =   1425
   End
   Begin VB.ComboBox cmbunit 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      ItemData        =   "FrmVeh_Master.frx":032F
      Left            =   5865
      List            =   "FrmVeh_Master.frx":0339
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox TxtResult 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   6945
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   45
      Width           =   1875
   End
   Begin VB.TextBox TxtHeight 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   4245
      TabIndex        =   2
      Top             =   60
      Width           =   1545
   End
   Begin VB.TextBox txtbreadth 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2370
      TabIndex        =   1
      Top             =   60
      Width           =   1545
   End
   Begin VB.TextBox TxtLength 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   435
      TabIndex        =   0
      Top             =   60
      Width           =   1545
   End
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
      Left            =   6375
      TabIndex        =   6
      Top             =   780
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
      Left            =   15
      TabIndex        =   5
      Top             =   990
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
      TabIndex        =   9
      Top             =   780
      Width           =   1065
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete Vehicle (Ctrl +D)"
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
      Left            =   9150
      TabIndex        =   8
      Top             =   780
      Width           =   1350
   End
   Begin VB.CommandButton CmdNew 
      Caption         =   "&New Vehicle (Ctrl +I)"
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
      Left            =   7755
      TabIndex        =   7
      Top             =   780
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   6915
      Left            =   45
      TabIndex        =   10
      Top             =   1350
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
         TabIndex        =   12
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   6735
         Left            =   30
         TabIndex        =   11
         Top             =   120
         Width           =   11550
         _ExtentX        =   20373
         _ExtentY        =   11880
         _Version        =   393216
         Rows            =   1
         Cols            =   5
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   3975
      TabIndex        =   16
      Top             =   150
      Width           =   210
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   2070
      TabIndex        =   15
      Top             =   135
      Width           =   210
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   150
      TabIndex        =   14
      Top             =   150
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Regn No. / Vehicle Name"
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
      Left            =   15
      TabIndex        =   13
      Top             =   750
      Width           =   2760
   End
End
Attribute VB_Name = "FrmVehMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbunit_Change()
    'Call unit_conversion
End Sub

Private Sub cmbunit_Click()
    Call unit_conversion
End Sub

Private Sub cmbunit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If cmbunit.ListIndex = -1 Then Exit Sub
            CmbUnit2.SetFocus
        Case vbKeyEscape
            TxtHeight.SetFocus
    End Select
End Sub

Private Sub CmbUnit2_Click()
    Call unit_conversion
End Sub

Private Sub CmbUnit2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            If cmbunit.ListIndex = -1 Then Exit Sub
            CmbUnit2.SetFocus
    End Select
End Sub

Private Sub CmdDelete_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    Dim i As Integer
    
    If GRDSTOCK.rows <= 1 Then Exit Sub
    On Error GoTo ErrHand
    
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) & " ****", vbYesNo + vbDefaultButton2, "DELETING ITEM....") = vbNo Then
        GRDSTOCK.SetFocus
        Exit Sub
    End If
    db.Execute ("DELETE from Veh_Master where Veh_Master.Veh_Code = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'")
    
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
    If GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 2) <> "" Then
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 0) = GRDSTOCK.rows - 1
        
        Dim TRXMAST As ADODB.Recordset
        On Error GoTo ErrHand
        
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select MAX(CONVERT(Veh_Code, SIGNED INTEGER)) From Veh_Master ", db, adOpenStatic, adLockReadOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            If IsNull(TRXMAST.Fields(0)) Then
                GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 1) = 1
            Else
                GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 1) = Val(TRXMAST.Fields(0)) + 1
            End If
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
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
    GRDSTOCK.TextMatrix(0, 1) = "" '"CODE"
    GRDSTOCK.TextMatrix(0, 2) = "REGN NO"
    GRDSTOCK.TextMatrix(0, 3) = "VEHICLE NAME"
    GRDSTOCK.TextMatrix(0, 4) = "VEHICLE DETAILS"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 1600
    GRDSTOCK.ColWidth(3) = 3500
    GRDSTOCK.ColWidth(4) = 5600
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 4
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    
    Call Fillgrid
    Me.Left = 500
    Me.Top = 0
    cmbunit.ListIndex = 0
    CmbUnit2.ListIndex = 1
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
                    Case 2
                        TXTsample.MaxLength = 20
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 3
                        If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)) = "" Then Exit Sub
                        TXTsample.MaxLength = 50
                        TXTsample.Visible = True
                        TXTsample.Top = GRDSTOCK.CellTop + 100
                        TXTsample.Left = GRDSTOCK.CellLeft '+ 60
                        TXTsample.Width = GRDSTOCK.CellWidth
                        TXTsample.Height = GRDSTOCK.CellHeight
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        TXTsample.SetFocus
                    Case 4
                        If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)) = "" Then Exit Sub
                        TXTsample.MaxLength = 100
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
    
    db.Execute "Update Veh_Master set veh_No ='' where isnull(veh_No)"
    db.Execute "Update Veh_Master set Veh_Name ='' where isnull(Veh_Name)"
    db.Execute "Update Veh_Master set Veh_Details ='' where isnull(Veh_Details)"
    GRDSTOCK.rows = 1
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM Veh_Master WHERE Veh_No Like '%" & tXTMEDICINE.text & "%' or Veh_Name Like '%" & tXTMEDICINE.text & "%' ORDER BY Veh_Name", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        'GRDSTOCK.FixedCols = 3
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!Veh_Code
        GRDSTOCK.TextMatrix(i, 2) = IIf(IsNull(rststock!veh_No), "", rststock!veh_No)
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!Veh_Name), "", rststock!Veh_Name)
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!Veh_Details), "", rststock!Veh_Details)
        
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

Private Sub txtbreadth_Change()
    Call unit_conversion
End Sub

Private Sub txtbreadth_GotFocus()
    txtbreadth.SelStart = 0
    txtbreadth.SelLength = Len(txtbreadth.text)
End Sub

Private Sub txtbreadth_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtbreadth.text) = 0 Then Exit Sub
            TxtHeight.SetFocus
        Case vbKeyEscape
            TxtLength.SetFocus
    End Select
End Sub

Private Sub txtbreadth_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtHeight_Change()
    Call unit_conversion
End Sub

Private Sub TxtHeight_GotFocus()
    TxtHeight.SelStart = 0
    TxtHeight.SelLength = Len(TxtHeight.text)
End Sub

Private Sub TxtHeight_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtHeight.text) = 0 Then Exit Sub
            cmbunit.SetFocus
        Case vbKeyEscape
            txtbreadth.SetFocus
    End Select
End Sub

Private Sub TxtHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtLength_Change()
    Call unit_conversion
End Sub

Private Sub TxtLength_GotFocus()
    TxtLength.SelStart = 0
    TxtLength.SelLength = Len(TxtLength.text)
End Sub

Private Sub TxtLength_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtLength.text) = 0 Then Exit Sub
            txtbreadth.SetFocus
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub TxtLength_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

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

Private Sub TxtResult_GotFocus()
    TxtResult.SelStart = 0
    TxtResult.SelLength = Len(TxtResult.text)
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
                Case 2 'reg No
                    If Trim(TXTsample.text) = "" Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from Veh_Master where Veh_Master.Veh_Code = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If (rststock.EOF And rststock.BOF) Then
                        rststock.AddNew
                        rststock!Veh_Code = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)
                        rststock!veh_No = Trim(TXTsample.text)
                    Else
                        rststock!veh_No = Trim(TXTsample.text)
                    End If
                    rststock.Update
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 3  'veh_name
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from Veh_Master where Veh_Master.Veh_Code = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!Veh_Name = Trim(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                                    
                Case 4  'veh details
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from Veh_Master where Veh_Master.Veh_Code = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!Veh_Details = Trim(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
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
        Case 2, 3
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
    End Select
End Sub


Private Function unit_conversion()
    If cmbunit.ListIndex = CmbUnit2.ListIndex Then
        TxtResult.text = Round(Val(TxtLength.text) * Val(txtbreadth.text) * Val(TxtHeight.text), 2)
    ElseIf cmbunit.ListIndex = 0 And CmbUnit2.ListIndex = 1 Then
        TxtResult.text = Round(35.3147 * Val(TxtLength.text) * Val(txtbreadth.text) * Val(TxtHeight.text), 0)
    ElseIf cmbunit.ListIndex = 1 And CmbUnit2.ListIndex = 0 Then
        TxtResult.text = Round(0.0283168 * Val(TxtLength.text) * Val(txtbreadth.text) * Val(TxtHeight.text), 2)
    End If
End Function
