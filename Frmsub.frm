VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMSUB 
   BackColor       =   &H0000FF00&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6420
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid grdsub 
      Height          =   3990
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   7038
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FRMSUB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim TMP As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPFLAG As Boolean
Dim M_STOCK As Integer

Private Sub Form_Load()
    Set grdsub.DataSource = Nothing
    TMP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & FRMSALE.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
    Set grdsub.DataSource = TMP
    grdsub.RowHeight = 250
    grdsub.Columns(0).Visible = False
    grdsub.Columns(1).Caption = "ITEM NAME"
    grdsub.Columns(1).Width = 4200
    grdsub.Columns(2).Caption = "QTY"
    grdsub.Columns(2).Width = 1300
    PHYFLAG = True
    FRMSALE.Enabled = False
    MDIMAIN.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHYFLAG = False Then PHY.Close
    TMP.Close
    FRMSALE.Enabled = True
    MDIMAIN.Enabled = True
End Sub

Private Sub grdsub_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            M_STOCK = Val(grdsub.Columns(2))
            If Trim(grdsub.Columns(2)) = "" Then Call STOCKADJUST
            If M_STOCK = 0 Then
                MsgBox "NO STOCK AVAILABLE..", vbOKOnly, "SALES"
                Exit Sub
            End If
            FRMSALE.TXTPRODUCT.Text = grdsub.Columns(1)
            FRMSALE.TXTITEMCODE.Text = grdsub.Columns(0)
            For i = 1 To FRMSALE.grdsales.Rows - 1
                If Trim(FRMSALE.grdsales.TextMatrix(i, 12)) = Trim(FRMSALE.TXTITEMCODE.Text) Then
                    If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
                        MDIMAIN.Enabled = True
                        FRMSALE.Enabled = True
                        FRMSALE.TXTPRODUCT.Enabled = True
                        FRMSALE.TxtQty.Enabled = False
                        FRMSALE.TXTPRODUCT.SetFocus
                        Unload Me
                        Exit Sub
                    End If
                End If
            Next i
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, SALES_PRICE, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            Set grdsub.DataSource = PHY
            If PHY.RecordCount = 1 Then
                FRMSALE.TxtQty.Text = grdsub.Columns(2)
                FRMSALE.TXTRATE.Text = grdsub.Columns(3)
                FRMSALE.TXTTAX.Text = grdsub.Columns(4)
                FRMSALE.txtexpdate.Text = grdsub.Columns(7)
                FRMSALE.txtBatch.Text = grdsub.Columns(6)
                
                FRMSALE.TXTVCHNO.Text = grdsub.Columns(8)
                FRMSALE.TXTLINENO.Text = grdsub.Columns(9)
                FRMSALE.TXTTRXTYPE.Text = grdsub.Columns(10)
                FRMSALE.txtunit.Text = grdsub.Columns(5)
                
                MDIMAIN.Enabled = True
                FRMSALE.Enabled = True
                FRMSALE.TXTPRODUCT.Enabled = False
                FRMSALE.TxtQty.Enabled = True
                FRMSALE.TxtQty.SetFocus
                Unload Me
                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                FRMSALE.Enabled = False
                FRMBATCH.Show
                Unload Me
            End If
        Case vbKeyEscape
            FRMSALE.TxtQty.Text = ""
            FRMSALE.TXTVCHNO.Text = ""
            FRMSALE.TXTLINENO.Text = ""
            FRMSALE.TXTTRXTYPE.Text = ""
            FRMSALE.txtunit.Text = ""
            MDIMAIN.Enabled = True
            FRMSALE.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = True
            FRMSALE.TxtQty.Enabled = False
            FRMSALE.TXTPRODUCT.SetFocus
            Unload Me
            Unload FRMBATCH
    End Select
End Sub

Private Function STOCKADJUST()
    Dim RSTSTOCK As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    M_STOCK = 0
    On Error GoTo ErrHand
    Set RSTSTOCK = New ADODB.Recordset
    RSTSTOCK.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & grdsub.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTSTOCK.EOF
        M_STOCK = M_STOCK + RSTSTOCK!BAL_QTY
        RSTSTOCK.MoveNext
    Loop
    RSTSTOCK.Close
    Set RSTSTOCK = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsub.Columns(0) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTITEMMAST
        If Not (.EOF And .BOF) Then
            !OPEN_QTY = M_STOCK
            !OPEN_VAL = 0
            !RCPT_QTY = 0
            !RCPT_VAL = 0
            !ISSUE_QTY = 0
            !ISSUE_VAL = 0
            !CLOSE_QTY = M_STOCK
            !CLOSE_VAL = 0
            !DAM_QTY = 0
            !DAM_VAL = 0
            RSTITEMMAST.Update
        End If
    End With
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Exit Function
    
ErrHand:
    MsgBox Err.Description
End Function
