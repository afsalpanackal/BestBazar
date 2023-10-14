VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMITEMLIST 
   BorderStyle     =   0  'None
   ClientHeight    =   5265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11115
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5265
   ScaleWidth      =   11115
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      Height          =   5100
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   11055
      Begin MSMask.MaskEdBox TXTEXPIRY 
         Height          =   300
         Left            =   375
         TabIndex        =   6
         Top             =   810
         Visible         =   0   'False
         Width           =   1080
         _ExtentX        =   1905
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##"
         PromptChar      =   " "
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
         Height          =   300
         Left            =   420
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4080
         Left            =   75
         TabIndex        =   1
         Top             =   645
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   7197
         _Version        =   393216
         Rows            =   1
         Cols            =   15
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DETAILS OF "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   1200
      End
      Begin VB.Label LBLPRODUCT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   1365
         TabIndex        =   3
         Top             =   285
         Width           =   4560
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PRESS ESC TO CANCEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   2
         Left            =   8850
         TabIndex        =   2
         Top             =   4800
         Width           =   2325
      End
   End
End
Attribute VB_Name = "FRMITEMLIST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "ITEM CODE"
    GRDBILL.TextMatrix(0, 2) = "ITEM NAME"
    GRDBILL.TextMatrix(0, 3) = "STOCKIST"
    GRDBILL.TextMatrix(0, 4) = "QTY"
    GRDBILL.TextMatrix(0, 5) = "" '"PACK"
    GRDBILL.TextMatrix(0, 6) = "Serial No"
    GRDBILL.TextMatrix(0, 7) = "EXPIRY"
    GRDBILL.TextMatrix(0, 8) = "MRP"
    GRDBILL.TextMatrix(0, 9) = "RATE"
    GRDBILL.TextMatrix(0, 10) = "G.PRICE"
    GRDBILL.TextMatrix(0, 13) = "INV NO."
    GRDBILL.TextMatrix(0, 14) = "INV DATE"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 0
    GRDBILL.ColWidth(2) = 0
    GRDBILL.ColWidth(3) = 2000
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 0 '700
    GRDBILL.ColWidth(6) = 900
    GRDBILL.ColWidth(7) = 1100
    GRDBILL.ColWidth(8) = 800
    GRDBILL.ColWidth(9) = 800
    GRDBILL.ColWidth(10) = 900
    GRDBILL.ColWidth(11) = 0
    GRDBILL.ColWidth(12) = 0
    GRDBILL.ColWidth(13) = 900
    GRDBILL.ColWidth(14) = 1100

    
    GRDBILL.ColAlignment(0) = 1
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 1
    GRDBILL.ColAlignment(3) = 1
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 1
    GRDBILL.ColAlignment(6) = 3
    GRDBILL.ColAlignment(7) = 3
    GRDBILL.ColAlignment(8) = 3
    GRDBILL.ColAlignment(9) = 3
    GRDBILL.ColAlignment(10) = 3
    GRDBILL.ColAlignment(13) = 3
    GRDBILL.ColAlignment(14) = 3
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Call FRMSTKSUMRY.fillstockgrid
End Sub

Private Sub GRDBILL_Click()
    TXTsample.Visible = False
    TXTEXPIRY.Visible = False
    GRDBILL.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If Not (GRDBILL.Col = 0 Or GRDBILL.Col = 1 Or GRDBILL.Col = 2 Or GRDBILL.Col = 3 Or GRDBILL.Col = 7 Or GRDBILL.Col = 13 Or GRDBILL.Col = 10 Or GRDBILL.Col = 14) Then
                TXTsample.Visible = True
                TXTsample.Top = GRDBILL.CellTop + 650
                TXTsample.Left = GRDBILL.CellLeft + 80
                TXTsample.Width = GRDBILL.CellWidth
                TXTsample.Text = GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col)
                TXTsample.SetFocus
            End If
            If GRDBILL.Col = 7 Then
                TXTEXPIRY.Visible = True
                TXTEXPIRY.Top = GRDBILL.CellTop + 650
                TXTEXPIRY.Left = GRDBILL.CellLeft + 80
                TXTEXPIRY.Width = GRDBILL.CellWidth
                TXTEXPIRY.Text = IIf(GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = "", "  /  ", Format(GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col), "MM/YY"))
                

                TXTEXPIRY.SetFocus
            End If
        Case vbKeyEscape
            FRMSTKSUMRy.Enabled = True
            Unload Me
            FRMSTKSUMRy.GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTEXPIRY.Text) = "/" Then
                M_DATE = "01/01/2001"
                GoTo SKIP
            End If
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
            
            M = Val(Mid(TXTEXPIRY.Text, 1, 2))
            Y = Val(Right(TXTEXPIRY.Text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            
SKIP:
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT RTRXFILE.EXP_DATE FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                rststock!EXP_DATE = IIf(M_DATE = "01/01/2001", Null, M_DATE)
                rststock.Update
            End If
            rststock.Close
            Set rststock = Nothing
            
            TXTEXPIRY.Visible = False
            GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = IIf(M_DATE = "01/01/2001", "", M_DATE)
            GRDBILL.Enabled = True
            GRDBILL.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDBILL.SetFocus
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDBILL.Col
                Case 4
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BAL_QTY = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                
                    M_STOCK = Val(TXTsample.Text) - Val(GRDBILL.TextMatrix(GRDBILL.Row, 4))
                    Set RSTITEMMAST = New ADODB.Recordset
                    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & GRDBILL.TextMatrix(GRDBILL.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    With RSTITEMMAST
                        If Not (.EOF And .BOF) Then
'                            !OPEN_QTY = M_STOCK
'                            !OPEN_VAL = 0
'                            !RCPT_QTY = 0
'                            !RCPT_VAL = 0
'                            !ISSUE_QTY = 0
'                            !ISSUE_VAL = 0
                            !CLOSE_QTY = !CLOSE_QTY + M_STOCK
'                            !CLOSE_VAL = 0
'                            !DAM_QTY = 0
'                            !DAM_VAL = 0
                            RSTITEMMAST.Update
                        End If
                    End With
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                    GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = TXTsample.Text
                    GRDBILL.Enabled = True
                    TXTsample.Visible = False
                Case 5
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!UNIT = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = TXTsample.Text
                    GRDBILL.Enabled = True
                    TXTsample.Visible = False
                Case 6
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REF_NO = Trim(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = TXTsample.Text
                    GRDBILL.Enabled = True
                    TXTsample.Visible = False
                Case 8
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = Format(Val(TXTsample.Text), ".000")
                    GRDBILL.Enabled = True
                    TXTsample.Visible = False
                Case 9
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.VCH_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 11)) & " AND RTRXFILE.LINE_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 12)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!SALES_PRICE = Val(TXTsample.Text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDBILL.TextMatrix(GRDBILL.Row, GRDBILL.Col) = Format(Val(TXTsample.Text), ".000")
                    GRDBILL.Enabled = True
                    TXTsample.Visible = False
            End Select
            GRDBILL.SetFocus
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDBILL.SetFocus
    End Select
End Sub


Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDBILL.Col
        Case 3, 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 5
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 7
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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

