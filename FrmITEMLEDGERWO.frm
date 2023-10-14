VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmStkAdjWO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK ADJUSTMENT"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6345
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmITEMLEDGERWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   6345
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
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6285
   End
   Begin VB.TextBox TXTsample 
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
      Height          =   290
      Left            =   3450
      TabIndex        =   4
      Top             =   3705
      Visible         =   0   'False
      Width           =   1350
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
      Height          =   405
      Left            =   4740
      TabIndex        =   3
      Top             =   7425
      Width           =   1290
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   5235
      Left            =   60
      TabIndex        =   2
      Top             =   2145
      Width           =   6270
      _ExtentX        =   11060
      _ExtentY        =   9234
      _Version        =   393216
      Rows            =   1
      Cols            =   12
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   8438015
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      Appearance      =   0
      GridLineWidth   =   2
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
   Begin MSDataListLib.DataList DataList2 
      Height          =   1425
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   2514
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
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   3105
      TabIndex        =   6
      Top             =   1890
      Width           =   3225
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
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
      Left            =   60
      TabIndex        =   5
      Top             =   1890
      Width           =   3045
   End
End
Attribute VB_Name = "FrmStkAdjWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REPFLAG As Boolean 'REP
Dim RSTREP As New ADODB.Recordset
Dim CLOSEALL As Integer

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            GRDSTOCK.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    REPFLAG = True
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "QTY"
    GRDSTOCK.TextMatrix(0, 4) = "" '"PACK"
    GRDSTOCK.TextMatrix(0, 5) = "Serial No"
    GRDSTOCK.TextMatrix(0, 6) = "EXPIRY"
    GRDSTOCK.TextMatrix(0, 7) = "MRP"
    GRDSTOCK.TextMatrix(0, 8) = "RATE"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 0
    GRDSTOCK.ColWidth(3) = 700
    GRDSTOCK.ColWidth(4) = 0 '750
    GRDSTOCK.ColWidth(5) = 1000
    GRDSTOCK.ColWidth(6) = 0
    GRDSTOCK.ColWidth(7) = 900
    GRDSTOCK.ColWidth(8) = 850
    GRDSTOCK.ColWidth(9) = 0
    GRDSTOCK.ColWidth(10) = 0
    GRDSTOCK.ColWidth(11) = 0
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 1
    GRDSTOCK.ColAlignment(6) = 4
    Me.Height = 8415
    Me.Width = 6465
    Me.Left = 2500
    Me.Top = 0
    CLOSEALL = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
     If CLOSEALL = 0 Then
        If REPFLAG = False Then RSTREP.Close
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
    End If
   Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3, 4, 5, 7
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 2150
                    TXTsample.Left = GRDSTOCK.CellLeft + 70
                    TXTsample.Width = GRDSTOCK.CellWidth - 25
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3  ' Bal QTY
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILEWO WHERE VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE= '" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BAL_QTY = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    M_STOCK = 0
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from [RTRXFILEWO] where ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db2, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        M_STOCK = M_STOCK + rststock!BAL_QTY
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
            
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT *  FROM ITEMMASTWO WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    With RSTITEMMAST
                        If Not (.EOF And .BOF) Then
                            '!OPEN_QTY = M_STOCK
                            '!OPEN_VAL = 0
                            '!RCPT_QTY = 0
                            '!RCPT_VAL = 0
                            '!ISSUE_QTY = 0
                            '!ISSUE_VAL = 0
                            !CLOSE_QTY = M_STOCK
                            '!CLOSE_VAL = 0
                            '!DAM_QTY = 0
                            '!DAM_VAL = 0
                            RSTITEMMAST.Update
                        End If
                    End With
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 4   'Pack
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILEWO WHERE VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE= '" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!UNIT = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                        rststock!SALES_PRICE = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) = Format(rststock!SALES_PRICE, "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 5  'BATCH
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILEWO WHERE RTRXFILEWO.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE= '" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REF_NO = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 7  'MRP
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILEWO WHERE VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE= '" & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(rststock!MRP, "0.000")
                        rststock!SALES_PRICE = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4)), 2)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) = Format(rststock!SALES_PRICE, "0.00")
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
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILEWO] where ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    
    Exit Function
    
eRRHAND:
    MsgBox Err.Description
End Function

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 4
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 5
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 7
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo eRRHAND
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"

    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            Call CMDEXIT_Click
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList2_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Long

    On Error GoTo eRRHAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GRDSTOCK.Rows = 1
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM RTRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY VCH_NO DESC", db2, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = rststock!BAL_QTY
        If IsNull(rststock!UNIT) Then
            GRDSTOCK.TextMatrix(i, 4) = 1
        Else
            GRDSTOCK.TextMatrix(i, 4) = rststock!UNIT
        End If
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDSTOCK.TextMatrix(i, 6) = ""
        GRDSTOCK.TextMatrix(i, 7) = Format(rststock!MRP, "0.000")
        GRDSTOCK.TextMatrix(i, 8) = Format(rststock!SALES_PRICE, "0.000")
        GRDSTOCK.TextMatrix(i, 9) = rststock!VCH_NO
        GRDSTOCK.TextMatrix(i, 10) = rststock!LINE_NO
        GRDSTOCK.TextMatrix(i, 11) = rststock!TRX_TYPE
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    LBLHEAD(9).Caption = "BATCH WISE LIST FOR THE ITEM "
    LBLHEAD(0).Caption = DataList2.Text
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
    
End Sub
