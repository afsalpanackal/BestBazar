VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSTKSUMRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK SUMMARY"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11250
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMSTKSUM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   11250
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
      Left            =   6195
      TabIndex        =   2
      Top             =   1560
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
      Height          =   450
      Left            =   10020
      TabIndex        =   1
      Top             =   8730
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   8625
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   15214
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   410
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblsvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   6810
      TabIndex        =   6
      Top             =   8745
      Width           =   1875
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sale Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   4485
      TabIndex        =   5
      Top             =   8850
      Width           =   2460
   End
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   2730
      TabIndex        =   4
      Top             =   8745
      Width           =   1920
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Purchase Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   90
      TabIndex        =   3
      Top             =   8850
      Width           =   2640
   End
End
Attribute VB_Name = "FRMSTKSUMRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim CLOSEALL As Integer

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Dim rststock As ADODB.Recordset
 
    Dim i As Integer
    Dim P_Value As Double
    Dim S_Value As Double
    
    'PHY_FLAG = True
    
    On Error GoTo eRRhAND
    CLOSEALL = 1
    i = 0
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "QTY"
    GRDSTOCK.TextMatrix(0, 4) = "" '"PACK"
    GRDSTOCK.TextMatrix(0, 5) = "Serial No"
    GRDSTOCK.TextMatrix(0, 6) = ""
    GRDSTOCK.TextMatrix(0, 7) = "MRP"
    GRDSTOCK.TextMatrix(0, 8) = "S.RATE"
    GRDSTOCK.TextMatrix(0, 11) = "COST"
    GRDSTOCK.TextMatrix(0, 12) = "Total P_Value"
    GRDSTOCK.TextMatrix(0, 13) = "Total S_Value"
    GRDSTOCK.TextMatrix(0, 14) = "TRX TYPE"
    
    GRDSTOCK.ColWidth(0) = 500
    GRDSTOCK.ColWidth(1) = 1500
    GRDSTOCK.ColWidth(2) = 2800
    GRDSTOCK.ColWidth(3) = 800
    GRDSTOCK.ColWidth(4) = 0 '800
    GRDSTOCK.ColWidth(5) = 0
    GRDSTOCK.ColWidth(6) = 0
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(8) = 900
    GRDSTOCK.ColWidth(9) = 0
    GRDSTOCK.ColWidth(10) = 0
    GRDSTOCK.ColWidth(11) = 900
    GRDSTOCK.ColWidth(12) = 1200
    GRDSTOCK.ColWidth(13) = 1200
    GRDSTOCK.ColWidth(14) = 0
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 1
    GRDSTOCK.ColAlignment(6) = 4
    
    Screen.MousePointer = vbHourglass
    
    S_Value = 0
    P_Value = 0
    'MDIMAIN.vbalProgressBar1.Visible = True
    'MDIMAIN.vbalProgressBar1.Value = 0
    'MDIMAIN.vbalProgressBar1.ShowText = True
    'MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM RTRXFILE WHERE  RTRXFILE.BAL_QTY > 0 ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic,adLockReadOnly
    rststock.Open "SELECT * FROM RTRXFILE ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        'MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = IIf(IsNull(rststock!ITEM_NAME), "", rststock!ITEM_NAME)
        GRDSTOCK.TextMatrix(i, 3) = Round(rststock!BAL_QTY, 2)
'        If IsNull(rststock!UNIT) Then
'            grdsTOCK.TextMatrix(i, 4) = 1
'        Else
'            grdsTOCK.TextMatrix(i, 4) = rststock!UNIT
'        End If
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
'        grdsTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
        GRDSTOCK.TextMatrix(i, 7) = Format(rststock!MRP, "0.000")
        GRDSTOCK.TextMatrix(i, 8) = Format(rststock!P_RETAIL, "0.000")
        GRDSTOCK.TextMatrix(i, 9) = rststock!VCH_NO
        GRDSTOCK.TextMatrix(i, 10) = rststock!LINE_NO
        GRDSTOCK.TextMatrix(i, 11) = Format(rststock!ITEM_COST, "0.000")
        GRDSTOCK.TextMatrix(i, 12) = Format(rststock!ITEM_COST * rststock!BAL_QTY, "0.00")
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 12)), "0.00")
        GRDSTOCK.TextMatrix(i, 13) = Format(rststock!P_RETAIL * rststock!BAL_QTY, "0.00")
        lblsvalue.Caption = Format(Val(lblsvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 13)), "0.00")
        GRDSTOCK.TextMatrix(i, 14) = rststock!TRX_TYPE
        rststock.MoveNext
        'MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
    Loop
    rststock.Close
    Set rststock = Nothing
    Me.Left = 500
    Me.Top = 0
    Me.Height = 10000
    Me.Width = 14595
    'MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        'If REPFLAG = False Then RSTREP.Close
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
    End If
   Cancel = CLOSEALL
'    MDIMAIN.PCTMENU.Enabled = True
'    'MDIMAIN.PCTMENU.Height = 555
'    MDIMAIN.PCTMENU.SetFocus
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
                ''Case 3, 4, 5, 7
                Case 3, 5
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 110
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth - 25
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Name...?", "STOCK"))
            For i = 1 To GRDSTOCK.Rows - 1
                    If Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem)) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
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
    
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3  ' Bal QTY
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BAL_QTY = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    M_STOCK = 0
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        M_STOCK = M_STOCK + rststock!BAL_QTY
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
            
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    With RSTITEMMAST
                        If Not (.EOF And .BOF) Then
'                            !OPEN_QTY = M_STOCK
'                            !OPEN_VAL = 0
'                            !RCPT_QTY = 0
'                            !RCPT_VAL = 0
'                            !ISSUE_QTY = 0
'                            !ISSUE_VAL = 0
                            !CLOSE_QTY = M_STOCK
'                            !CLOSE_VAL = 0
'                            !DAM_QTY = 0
'                            !DAM_VAL = 0
                            RSTITEMMAST.Update
                        End If
                    End With
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    
                Case 4   'Pack
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!UNIT = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                        rststock!P_RETAIL = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) / Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) = Format(rststock!P_RETAIL, "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format(rststock!P_RETAIL * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    
                Case 5  'BATCH
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REF_NO = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.Text
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    
                Case 7  'MRP
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.Text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(rststock!MRP, "0.000")
                        ''Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * 15 / 100, ".000")
                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4)), 2)
                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) - Val(TXTsample.Text) * 15 / 100, 2)
                        ''grdsTOCK.TextMatrix(grdsTOCK.Row, 8) = Format(rststock!P_RETAIL, "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format(rststock!P_RETAIL * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False

            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        M_STOCK = M_STOCK + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    
    Exit Function
    
eRRhAND:
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

Private Function TOTALVALUE()
    Dim i As Integer
    
    lblsvalue.Caption = ""
    lblpvalue.Caption = ""
    For i = 1 To GRDSTOCK.Rows - 1
        lblsvalue.Caption = Val(lblsvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 13))
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 12)), "0.00")
    Next i
    lblsvalue.Caption = Format(lblsvalue.Caption, "0.00")
    lblpvalue.Caption = Format(lblpvalue.Caption, "0.00")
    
End Function
