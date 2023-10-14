VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmDead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Non Moving Items"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   11580
   Icon            =   "FrmDead.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   11580
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export"
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
      Left            =   15
      TabIndex        =   7
      Top             =   7920
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00B4D6DA&
      Height          =   645
      Left            =   15
      TabIndex        =   3
      Top             =   -75
      Width           =   5760
      Begin VB.OptionButton Option1 
         BackColor       =   &H00B4D6DA&
         Caption         =   "&Not Sold After"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   2430
         TabIndex        =   6
         Top             =   225
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton OptFull 
         BackColor       =   &H00B4D6DA&
         Caption         =   "&Not Sold Yet"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   270
         Left            =   135
         TabIndex        =   4
         Top             =   225
         Width           =   2145
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   4155
         TabIndex        =   5
         Top             =   180
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   192
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   111214593
         CurrentDate     =   40498
      End
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
      Left            =   10335
      TabIndex        =   1
      Top             =   7905
      Width           =   1200
   End
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
      Left            =   9000
      TabIndex        =   0
      Top             =   7905
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   7290
      Left            =   0
      TabIndex        =   2
      Top             =   570
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   12859
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
      AllowUserResizing=   3
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
Attribute VB_Name = "FrmDead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDDISPLAY_Click()
    Dim rststock, RSTRTRXFILE, RSTSUPPLIER, RSTSUPPLIER2 As ADODB.Recordset
    Dim i As Long
    
    'PHY_FLAG = True
    
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    'Screen.MousePointer = vbHourglass
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenForwardOnly
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM ITEMMAST WHERE CLOSE_QTY >0 ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        Set RSTRTRXFILE = New ADODB.Recordset
        If OptFull.value = True Then
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Else
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
        End If
        If (RSTRTRXFILE.EOF Or RSTRTRXFILE.BOF) Then
            i = i + 1
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            GRDSTOCK.FixedRows = 1
            GRDSTOCK.TextMatrix(i, 0) = i
            GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
            GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
            GRDSTOCK.TextMatrix(i, 3) = "Opening Stock"
            GRDSTOCK.TextMatrix(i, 4) = rststock!CLOSE_QTY
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                GRDSTOCK.TextMatrix(i, 3) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
            Else
                Set RSTSUPPLIER2 = New ADODB.Recordset
                RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                    GRDSTOCK.TextMatrix(i, 3) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                End If
                RSTSUPPLIER2.Close
                Set RSTSUPPLIER2 = Nothing
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
        End If
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
SKIP:
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
        
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Dead Stock Items"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, N As Long
    
    On Error GoTo eRRhAND
    Screen.MousePointer = vbHourglass
    'Create an Excel instalce.
    Set oApp = CreateObject("Excel.Application")
    Set oWB = oApp.Workbooks.Add
    Set oWS = oWB.Worksheets(1)
    

    
    
'    xlRange = oWS.Range("A1", "C1")
'    xlRange.Font.Bold = True
'    xlRange.ColumnWidth = 15
'    'xlRange.Value = {"First Name", "Last Name", "Last Service"}
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
'
'    xlRange = oWS.Range("C1", "C999")
'    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'    xlRange.ColumnWidth = 12
    
    'If Sum_flag = False Then
        oWS.Range("A1", "E1").Merge
        oWS.Range("A1", "E1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "E2").Merge
        oWS.Range("A2", "E2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    
    
    oWS.Range("A1").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 14            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    'oApp.Columns("A:A").EntireColumn.AutoFit     '-- autofitted column

    oWS.Range("A2").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 11            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True

    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).value = "DEAD STOCK REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).value = GRDSTOCK.TextMatrix(0, 4)
    
    On Error GoTo eRRhAND
    
    i = 4
    For N = 1 To GRDSTOCK.Rows - 1
        oWS.Range("A" & i).value = GRDSTOCK.TextMatrix(N, 0)
        oWS.Range("B" & i).value = GRDSTOCK.TextMatrix(N, 1)
        oWS.Range("C" & i).value = GRDSTOCK.TextMatrix(N, 2)
        oWS.Range("D" & i).value = GRDSTOCK.TextMatrix(N, 3)
        oWS.Range("E" & i).value = GRDSTOCK.TextMatrix(N, 4)
        On Error GoTo eRRhAND
        i = i + 1
    Next N
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
   
SKIP:
    oApp.Visible = True
    
    
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "LAST SUPPLIER"
    GRDSTOCK.TextMatrix(0, 4) = "BAL QTY"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 4200
    GRDSTOCK.ColWidth(3) = 4700
    GRDSTOCK.ColWidth(4) = 1300
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    Left = 500
    Top = 0
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
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

