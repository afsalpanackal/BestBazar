VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMZEROSTOCK 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK SUMMARY"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMZEROSTOCK.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   13200
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display"
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
      Left            =   10365
      TabIndex        =   2
      Top             =   8400
      Width           =   1200
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
      Left            =   11700
      TabIndex        =   1
      Top             =   8400
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   8340
      Left            =   30
      TabIndex        =   0
      Top             =   15
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   14711
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
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
End
Attribute VB_Name = "FRMZEROSTOCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CLOSEALL As Integer
Dim MFG_REC As New ADODB.Recordset
Dim SCH_REC As New ADODB.Recordset
Dim MOLEFLAG As Boolean
Dim RSTMOLE As New ADODB.Recordset

Private Sub CMDDISPLAY_Click()
    Dim rststock, RSTRTRXFILE, RSTSUPPLIER As ADODB.Recordset
    Dim i As Integer
    
    'PHY_FLAG = True
    
    
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    
    Screen.MousePointer = vbHourglass
    On Error GoTo errHand
    Set rststock = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenForwardOnly
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME FROM RTRXFILE ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.Value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND CLOSE_QTY <= 0 ", db, adOpenStatic, adLockReadOnly
        If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
            i = i + 1
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            GRDSTOCK.FixedRows = 1
            GRDSTOCK.TextMatrix(i, 0) = i
            GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
            GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                Select Case RSTSUPPLIER!TRX_TYPE
                    Case "PI"
                        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", Mid(RSTSUPPLIER!VCH_DESC, 15))
                End Select
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
            GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(RSTRTRXFILE!SCHEDULE), "", RSTRTRXFILE!SCHEDULE)
            
            rststock.MoveNext
        End If
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
        
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

errHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Form_Load()

    CLOSEALL = 1
    
    MOLEFLAG = True
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "SUPPLIER"
    GRDSTOCK.TextMatrix(0, 4) = "SCHEDULE"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 4200
    GRDSTOCK.ColWidth(3) = 5000
    GRDSTOCK.ColWidth(4) = 1000
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    
    
    Left = 500
    Top = 0
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        'If REPFLAG = False Then RSTREP.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 555
        MDIMAIN.PCTMENU.SetFocus
    End If
   Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 114
            sitem = UCase(InputBox("Item Name..?", "ZERO STOCK"))
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
