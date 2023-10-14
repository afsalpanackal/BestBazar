VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmStkmovmntwo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK MOVEMENT"
   ClientHeight    =   9510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmstkmvmtwo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9510
   ScaleWidth      =   18465
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SORT ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1470
      Left            =   11190
      TabIndex        =   19
      Top             =   15
      Width           =   7245
      Begin VB.OptionButton OPTOUTDATE 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   105
         TabIndex        =   21
         Top             =   360
         Value           =   -1  'True
         Width           =   3285
      End
      Begin VB.OptionButton OPTOUTCUST 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   720
         Width           =   3285
      End
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
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   6195
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
      Left            =   13470
      TabIndex        =   3
      Top             =   9030
      Width           =   1380
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   7155
      Left            =   45
      TabIndex        =   2
      Top             =   1770
      Width           =   10995
      _ExtentX        =   19394
      _ExtentY        =   12621
      _Version        =   393216
      Rows            =   1
      Cols            =   15
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
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1035
      Left            =   45
      TabIndex        =   1
      Top             =   435
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1826
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SORT ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1485
      Left            =   6270
      TabIndex        =   6
      Top             =   15
      Width           =   4770
      Begin VB.OptionButton OPTBALQTY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Available Qty"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   135
         TabIndex        =   9
         Top             =   900
         Width           =   3285
      End
      Begin VB.OptionButton optsupplier 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Supplier"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   570
         Width           =   3285
      End
      Begin VB.OptionButton optdate 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   105
         TabIndex        =   7
         Top             =   225
         Value           =   -1  'True
         Width           =   3285
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GRDOUTWARD 
      Height          =   7155
      Left            =   11190
      TabIndex        =   14
      Top             =   1770
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   12621
      _Version        =   393216
      Rows            =   1
      Cols            =   9
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
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   4  'Dash-Dot
      BorderWidth     =   2
      DrawMode        =   4  'Mask Not Pen
      X1              =   11115
      X2              =   11100
      Y1              =   30
      Y2              =   8955
   End
   Begin VB.Label lblmanual 
      Caption         =   "*Adjusted Manually"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   18
      Top             =   9060
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
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
      Height          =   225
      Index           =   1
      Left            =   11190
      TabIndex        =   17
      Top             =   1515
      Width           =   7230
   End
   Begin VB.Label LBLOUTWARD 
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   5460
      TabIndex        =   16
      Top             =   9060
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AVAILABLE QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   1
      Left            =   7155
      TabIndex        =   15
      Top             =   9060
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OUTWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   0
      Left            =   3420
      TabIndex        =   13
      Top             =   9060
      Width           =   1785
   End
   Begin VB.Label LBLBALANCE 
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
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   9195
      TabIndex        =   12
      Top             =   9060
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   9060
      Width           =   1545
   End
   Begin VB.Label LBLINWARD 
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
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   1800
      TabIndex        =   10
      Top             =   9060
      Width           =   1500
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
      Left            =   3090
      TabIndex        =   5
      Top             =   1515
      Width           =   7950
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
      Left            =   45
      TabIndex        =   4
      Top             =   1515
      Width           =   3750
   End
End
Attribute VB_Name = "FrmStkmovmntwo"
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
    GRDSTOCK.TextMatrix(0, 3) = "TYPE"
    GRDSTOCK.TextMatrix(0, 4) = "SUPPLIER"
    GRDSTOCK.TextMatrix(0, 5) = "QTY"
    GRDSTOCK.TextMatrix(0, 6) = "INV DATE"
    GRDSTOCK.TextMatrix(0, 7) = "INV NO"
    GRDSTOCK.TextMatrix(0, 8) = "COMP REF"
    GRDSTOCK.TextMatrix(0, 9) = "" '"PACK"
    GRDSTOCK.TextMatrix(0, 10) = "Serial No"
    GRDSTOCK.TextMatrix(0, 11) = "EXPIRY"
    GRDSTOCK.TextMatrix(0, 12) = "MRP"
    GRDSTOCK.TextMatrix(0, 13) = "G. RATE"
    GRDSTOCK.TextMatrix(0, 14) = "BAL QTY"
    
    GRDSTOCK.ColWidth(0) = 400
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 0
    GRDSTOCK.ColWidth(3) = 1000
    GRDSTOCK.ColWidth(4) = 2050
    GRDSTOCK.ColWidth(5) = 1200
    GRDSTOCK.ColWidth(6) = 1200
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(8) = 900
    GRDSTOCK.ColWidth(9) = 0 '700
    GRDSTOCK.ColWidth(10) = 900
    GRDSTOCK.ColWidth(11) = 0
    GRDSTOCK.ColWidth(12) = 0
    GRDSTOCK.ColWidth(13) = 700
    GRDSTOCK.ColWidth(14) = 800

    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 1
     GRDSTOCK.ColAlignment(8) = 1
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 1
    GRDSTOCK.ColAlignment(11) = 4
    GRDSTOCK.ColAlignment(12) = 1
    GRDSTOCK.ColAlignment(13) = 4
    GRDSTOCK.ColAlignment(14) = 4
    
    GRDOUTWARD.TextMatrix(0, 0) = "SL"
    GRDOUTWARD.TextMatrix(0, 1) = "TYPE"
    GRDOUTWARD.TextMatrix(0, 2) = "CUSTOMER"
    GRDOUTWARD.TextMatrix(0, 3) = "QTY"
    GRDOUTWARD.TextMatrix(0, 4) = "FREE"
    GRDOUTWARD.TextMatrix(0, 5) = "RATE"
    GRDOUTWARD.TextMatrix(0, 6) = "INV #"
    GRDOUTWARD.TextMatrix(0, 7) = "INV DATE"
    GRDOUTWARD.TextMatrix(0, 8) = "Serial No"

    GRDOUTWARD.ColWidth(0) = 400
    GRDOUTWARD.ColWidth(1) = 1100
    GRDOUTWARD.ColWidth(2) = 2100
    GRDOUTWARD.ColWidth(3) = 800
    GRDOUTWARD.ColWidth(4) = 800
    GRDOUTWARD.ColWidth(5) = 800
    GRDOUTWARD.ColWidth(6) = 900
    GRDOUTWARD.ColWidth(7) = 1200
    GRDOUTWARD.ColWidth(8) = 1200
    
    GRDOUTWARD.ColAlignment(0) = 4
    GRDOUTWARD.ColAlignment(1) = 1
    GRDOUTWARD.ColAlignment(2) = 1
    GRDOUTWARD.ColAlignment(3) = 4
    GRDOUTWARD.ColAlignment(4) = 4
    GRDOUTWARD.ColAlignment(5) = 4
    GRDOUTWARD.ColAlignment(6) = 1
    GRDOUTWARD.ColAlignment(7) = 4
    GRDOUTWARD.ColAlignment(8) = 4
  
    Me.Height = 9990
    Me.Width = 18555
    Me.Left = 0
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

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            
    End Select
End Sub

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILEWO] where RTRXFILEWO.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db2, adOpenStatic, adLockReadOnly
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

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub OPTBALQTY_Click()
    Call fillgrid
    Call FILLGRID2
End Sub

Private Sub optdate_Click()
    Call fillgrid
    Call FILLGRID2
End Sub

Private Sub OPTOUTCUST_Click()
    Call fillgrid
    Call FILLGRID2
End Sub

Private Sub OPTOUTDATE_Click()
    Call fillgrid
    Call FILLGRID2
End Sub

Private Sub optsupplier_Click()
    Call fillgrid
    Call FILLGRID2
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo eRRHAND
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT [ITEM_CODE], [ITEM_NAME] From ITEMMASTWO  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY [ITEM_NAME]", db2, adOpenStatic, adLockReadOnly
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
    Call fillgrid
    Call FILLGRID2
    ''''''''LBLBALANCE.Caption = Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption)
End Sub

Private Function fillgrid()
    Dim rststock As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo eRRHAND
    
    i = 0
    Screen.MousePointer = vbHourglass
    LBLINWARD.Caption = ""
    LBLBALANCE.Caption = ""
    
    GRDSTOCK.Rows = 1
    Set rststock = New ADODB.Recordset
    If optdate.Value = True Then rststock.Open "SELECT * FROM RTRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY VCH_DATE DESC", db2, adOpenStatic, adLockReadOnly
    If optsupplier.Value = True Then rststock.Open "SELECT * FROM RTRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY TRX_TYPE,VCH_DESC", db2, adOpenStatic, adLockReadOnly
    If OPTBALQTY.Value = True Then rststock.Open "SELECT * FROM RTRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "'ORDER BY BAL_QTY DESC", db2, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        Select Case rststock!TRX_TYPE
            Case "CN"
                GRDSTOCK.TextMatrix(i, 3) = "SALES RETURN"
                GRDSTOCK.TextMatrix(i, 4) = Mid(rststock!VCH_DESC, 15)
            Case "XX", "OP"
                GRDSTOCK.TextMatrix(i, 3) = "OPENING STOCK"
            Case Else
                GRDSTOCK.TextMatrix(i, 3) = "PURCHASE"
                GRDSTOCK.TextMatrix(i, 4) = Mid(rststock!VCH_DESC, 15)
        End Select
        GRDSTOCK.TextMatrix(i, 5) = rststock!QTY ''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
        'GRDSTOCK.TextMatrix(i, 3) = rststock!BAL_QTY
        GRDSTOCK.TextMatrix(i, 6) = rststock!VCH_DATE
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!PINV), "", rststock!PINV)
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!PINV), "", rststock!VCH_NO)
        If IsNull(rststock!UNIT) Then
            GRDSTOCK.TextMatrix(i, 9) = 1
        Else
            GRDSTOCK.TextMatrix(i, 9) = rststock!UNIT
        End If
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
        GRDSTOCK.TextMatrix(i, 11) = ""
        GRDSTOCK.TextMatrix(i, 12) = Format(rststock!MRP, "0.000")
        GRDSTOCK.TextMatrix(i, 13) = Format(rststock!ITEM_COST, "0.000")
        GRDSTOCK.TextMatrix(i, 14) = rststock!BAL_QTY
        LBLINWARD.Caption = Val(LBLINWARD.Caption) + rststock!QTY '''(rststock!QTY / rststock!UNIT) * rststock!LINE_DISC
        LBLBALANCE.Caption = Val(LBLBALANCE.Caption) + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    LBLHEAD(9).Caption = "INWARD DETAILS FOR THE ITEM "
    LBLHEAD(0).Caption = DataList2.Text
    Screen.MousePointer = vbNormal
    Exit Function

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Function FILLGRID2()
    Dim rststock As ADODB.Recordset
    Dim RSTTEMP As ADODB.Recordset
    Dim M As Integer
    Dim E_DATE As Date
    Dim i, OUTQTY As Long

    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo eRRHAND
    
    Screen.MousePointer = vbHourglass
    LBLOUTWARD.Caption = ""
    
    db2.Execute "delete * From TEMPTRX"
    Set RSTTEMP = New ADODB.Recordset
    RSTTEMP.Open "SELECT *  FROM TEMPTRX", db2, adOpenStatic, adLockOptimistic, adCmdText
    OUTQTY = 0
    i = 0
    GRDOUTWARD.Rows = 1
    Set rststock = New ADODB.Recordset
    If OPTOUTDATE.Value = True Then rststock.Open "SELECT * FROM TRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='PR' OR TRX_TYPE='DG' OR TRX_TYPE='GF' OR TRX_TYPE='SR') ORDER BY VCH_DATE DESC", db2, adOpenStatic, adLockReadOnly
    If OPTOUTCUST.Value = True Then rststock.Open "SELECT * FROM TRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='PR' OR TRX_TYPE='DG' OR TRX_TYPE='GF' OR TRX_TYPE='SR') ORDER BY VCH_DESC ASC, VCH_DATE DESC", db2, adOpenStatic, adLockReadOnly
    'rststock.Open "SELECT * FROM TRXFILEWO WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND CST <>2", DB2, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDOUTWARD.Rows = GRDOUTWARD.Rows + 1
        GRDOUTWARD.FixedRows = 1
        GRDOUTWARD.TextMatrix(i, 0) = i
'        Select Case rststock!CST
'            Case 0
'               GRDOUTWARD.TextMatrix(i, 1) = "SALES"
'            Case 1
'               GRDOUTWARD.TextMatrix(i, 1) = "DELIVEREY"
'            Case 2
'               GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
'        End Select
        GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
        Select Case rststock!TRX_TYPE
            Case "SI"
                GRDOUTWARD.TextMatrix(i, 1) = "WHOLESALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "RI"
                GRDOUTWARD.TextMatrix(i, 1) = "RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DN"
                GRDOUTWARD.TextMatrix(i, 1) = "DELIVERY"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "PR"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DG"
                GRDOUTWARD.TextMatrix(i, 1) = "DAMAGED GOODS"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GF"
                GRDOUTWARD.TextMatrix(i, 1) = "SAMPLE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "SR"
                GRDOUTWARD.TextMatrix(i, 1) = "TO SERVICE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
        End Select
        
        
        GRDOUTWARD.TextMatrix(i, 3) = rststock!QTY
        GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), "", rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 5) = Format(rststock!SALES_PRICE, "0.00")
        GRDOUTWARD.TextMatrix(i, 6) = rststock!TRX_TYPE & "-" & rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 7) = Format(rststock!VCH_DATE, "dd/mm/yyyy")
        GRDOUTWARD.TextMatrix(i, 8) = rststock!REF_NO
        LBLOUTWARD.Caption = Val(LBLOUTWARD.Caption) + rststock!QTY + Val(GRDOUTWARD.TextMatrix(i, 4))
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
        
    LBLHEAD(1).Caption = "OUTWARD DETAILS"
    Screen.MousePointer = vbNormal
    If Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption) <> Val(LBLBALANCE.Caption) Then lblmanual.Visible = True Else lblmanual.Visible = False
    Exit Function

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

