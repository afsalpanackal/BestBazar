VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FRMSTKSUMRY 
   Caption         =   "STOCK SUMMARY"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMSTKSUM.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   7755
   Begin VB.Frame FRMEMAIN 
      Height          =   8580
      Left            =   30
      TabIndex        =   3
      Top             =   -30
      Width           =   7710
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   8385
         Left            =   -15
         TabIndex        =   0
         Top             =   120
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   14790
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
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
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&PRINT"
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
      Left            =   4395
      TabIndex        =   1
      Top             =   8700
      Width           =   1215
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
      Left            =   5715
      TabIndex        =   2
      Top             =   8700
      Width           =   1185
   End
   Begin Crystal.CrystalReport rptprint 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "FRMSTKSUMRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub cmdprint_Click()
    rptPRINT.ReportFileName = App.Path & "\RPTSTOCK.RPT"
    rptPRINT.Action = 1
End Sub

Private Sub Form_Load()
        
    Me.Height = 10000
    Me.Width = 7875
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = (Screen.Height - Me.Height) - 3000 / 2
    Call fillstockgrid

End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim rststock As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDSTOCK.Rows = 1 Then Exit Sub
            FRMITEMLIST.GRDBILL.Rows = 1
            i = 0
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' ORDER BY RTRXFILE.VCH_NO", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until rststock.EOF
                i = i + 1
                FRMITEMLIST.GRDBILL.Rows = FRMITEMLIST.GRDBILL.Rows + 1
                FRMITEMLIST.GRDBILL.FixedRows = 1
                FRMITEMLIST.GRDBILL.TextMatrix(i, 0) = i
                FRMITEMLIST.GRDBILL.TextMatrix(i, 1) = rststock!ITEM_CODE
                FRMITEMLIST.GRDBILL.TextMatrix(i, 2) = rststock!ITEM_NAME
                Set rststockist = New ADODB.Recordset
                rststockist.Open "SELECT ACT_NAME FROM ACTMAST WHERE ACT_CODE = '" & rststock!M_USER_ID & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (rststockist.EOF And rststockist.BOF) Then
                    FRMITEMLIST.GRDBILL.TextMatrix(i, 3) = rststockist!ACT_NAME
                End If
                rststockist.Close
                Set rststockist = Nothing
                FRMITEMLIST.GRDBILL.TextMatrix(i, 4) = rststock!BAL_QTY
                If IsNull(rststock!UNIT) Then
                    FRMITEMLIST.GRDBILL.TextMatrix(i, 5) = 1
                Else
                    FRMITEMLIST.GRDBILL.TextMatrix(i, 5) = rststock!UNIT
                End If
                FRMITEMLIST.GRDBILL.TextMatrix(i, 6) = rststock!REF_NO
                FRMITEMLIST.GRDBILL.TextMatrix(i, 7) = IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
                FRMITEMLIST.GRDBILL.TextMatrix(i, 8) = Format(rststock!MRP, ".000")
                FRMITEMLIST.GRDBILL.TextMatrix(i, 9) = Format(rststock!SALES_PRICE, ".000")
                FRMITEMLIST.GRDBILL.TextMatrix(i, 10) = Format(rststock!ITEM_COST, ".000")
                FRMITEMLIST.GRDBILL.TextMatrix(i, 11) = rststock!VCH_NO
                FRMITEMLIST.GRDBILL.TextMatrix(i, 12) = rststock!LINE_NO
                FRMITEMLIST.GRDBILL.TextMatrix(i, 13) = rststock!PINV
                FRMITEMLIST.GRDBILL.TextMatrix(i, 14) = rststock!VCH_DATE
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            FRMITEMLIST.LBLPRODUCT.Caption = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)
            FRMSTKSUMRY.Enabled = False
            FRMITEMLIST.Show
            FRMITEMLIST.GRDBILL.SetFocus
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

Public Function fillstockgrid()
    Dim rststock As ADODB.Recordset
    Dim i As Integer
    Dim N As Integer
    
    
    i = 0
    N = GRDSTOCK.Row
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "QTY"
    
    
    GRDSTOCK.ColWidth(0) = 500
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 5000
    GRDSTOCK.ColWidth(3) = 1000
    GRDSTOCK.ColAlignment(3) = 3
   '
    Screen.MousePointer = vbHourglass

    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM ITEMMAST ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    GRDSTOCK.Rows = 1
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = rststock!CLOSE_QTY
    
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    GRDSTOCK.TopRow = IIf(N = 0, 1, N)
    GRDSTOCK.Row = N
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
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


