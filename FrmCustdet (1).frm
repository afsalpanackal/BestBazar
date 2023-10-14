VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCustList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CUSTOMER DETAILS"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18945
   Icon            =   "FrmCustdet.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   18945
   Begin VB.CommandButton Command2 
      Caption         =   "Display All Cash Customers List"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   11565
      TabIndex        =   22
      Top             =   7860
      Width           =   1575
   End
   Begin VB.CommandButton CmdCashCust 
      Caption         =   "Cash Customers List with Phone"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9690
      TabIndex        =   21
      Top             =   7860
      Width           =   1845
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   13860
      TabIndex        =   17
      Top             =   -60
      Width           =   5100
      Begin VB.OptionButton OptAllAgents 
         BackColor       =   &H00C0E0FF&
         Caption         =   "All"
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
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton OptAgent 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Agent"
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
         Height          =   315
         Left            =   885
         TabIndex        =   18
         Top             =   180
         Width           =   870
      End
      Begin MSForms.ComboBox CmbAgent 
         Height          =   360
         Left            =   1755
         TabIndex        =   20
         Top             =   165
         Width           =   3300
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "5821;635"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   255
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export to Excel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   6165
      TabIndex        =   16
      Top             =   7875
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00F4F0DB&
      Height          =   615
      Left            =   9120
      TabIndex        =   13
      Top             =   -60
      Width           =   4740
      Begin VB.OptionButton OptSupplier 
         BackColor       =   &H00F4F0DB&
         Caption         =   "Suppliers (Credtors)"
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
         Height          =   315
         Left            =   2385
         TabIndex        =   15
         Top             =   180
         Width           =   2175
      End
      Begin VB.OptionButton Optcustomers 
         BackColor       =   &H00F4F0DB&
         Caption         =   "Customers (Debtors)"
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
         Height          =   315
         Left            =   75
         TabIndex        =   14
         Top             =   180
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
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
      Left            =   45
      TabIndex        =   12
      Top             =   7890
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7365
      TabIndex        =   6
      Top             =   7860
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   615
      Left            =   45
      TabIndex        =   2
      Top             =   -60
      Width           =   9075
      Begin VB.TextBox TxtName 
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
         Height          =   360
         Left            =   1380
         TabIndex        =   8
         Top             =   165
         Width           =   2745
      End
      Begin VB.TextBox txtCode 
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
         Height          =   360
         Left            =   45
         TabIndex        =   7
         Top             =   165
         Width           =   1320
      End
      Begin VB.OptionButton optCategory 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Area"
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
         Height          =   315
         Left            =   4815
         TabIndex        =   5
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton OptAllCategory 
         BackColor       =   &H00C0E0FF&
         Caption         =   "All"
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
         Height          =   315
         Left            =   4155
         TabIndex        =   4
         Top             =   180
         Value           =   -1  'True
         Width           =   795
      End
      Begin MSForms.ComboBox cmbarea 
         Height          =   360
         Left            =   5685
         TabIndex        =   3
         Top             =   165
         Width           =   3300
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "5821;635"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         BorderColor     =   255
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
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
      Height          =   495
      Left            =   13200
      TabIndex        =   1
      Top             =   7860
      Width           =   1125
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
      Height          =   495
      Left            =   8505
      TabIndex        =   0
      Top             =   7860
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   7350
      Left            =   0
      TabIndex        =   9
      Top             =   465
      Width           =   18975
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
         Height          =   405
         Left            =   7470
         TabIndex        =   11
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   7245
         Left            =   15
         TabIndex        =   10
         Top             =   90
         Width           =   18945
         _ExtentX        =   33417
         _ExtentY        =   12779
         _Version        =   393216
         Rows            =   1
         Cols            =   8
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
End
Attribute VB_Name = "FrmCustList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Fillgrid()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "CODE"
    GRDTranx.TextMatrix(0, 2) = "NAME"
    GRDTranx.TextMatrix(0, 3) = "ADDRESS"
    GRDTranx.TextMatrix(0, 4) = "PHONE"
    GRDTranx.TextMatrix(0, 5) = "GST NO."
    GRDTranx.TextMatrix(0, 6) = "AREA"
    GRDTranx.TextMatrix(0, 7) = "OP BAL"
        
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 1000
    GRDTranx.ColWidth(2) = 4000
    GRDTranx.ColWidth(3) = 6000
    GRDTranx.ColWidth(4) = 1700
    GRDTranx.ColWidth(5) = 1700
    GRDTranx.ColWidth(6) = 1700
    GRDTranx.ColWidth(7) = 1200
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    
    Dim rstTRANX, rstCust As ADODB.Recordset
    Dim i As Long
    
    If optCategory.value = True And cmbarea.Text = "" Then
        MsgBox "Please select the Place from the List", vbOKOnly, "Receipt Dues"
        On Error Resume Next
        cmbarea.SetFocus
        Exit Function
    End If
    If Optcustomers.value = True And OptAgent.value = True And CmbAgent.Text = "" Then
        MsgBox "Please select the Agent from the List", vbOKOnly, "Receipt Dues"
        On Error Resume Next
        CmbAgent.SetFocus
        Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    i = 1
    On Error GoTo eRRHAND
    
    Set rstCust = New ADODB.Recordset
    If Optcustomers.value = True Then
        GRDTranx.Cols = 9
        GRDTranx.TextMatrix(0, 8) = "Agent"
        GRDTranx.ColWidth(8) = 1500
        GRDTranx.ColAlignment(8) = 1
        
        If optCategory.value = True Then
            If OptAgent.value = True Then
                rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' AND AGENT_NAME Like '%" & CmbAgent.Text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
            Else
                rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
            End If
        Else
            If OptAgent.value = True Then
                rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' AND AGENT_NAME Like '%" & CmbAgent.Text & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
            Else
                rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
            End If
        End If
        Do Until rstCust.EOF
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(i, 0) = i
            GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!ACT_CODE), "", rstCust!ACT_CODE)
            GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!ACT_NAME), "", rstCust!ACT_NAME)
            GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstCust!Address), "", rstCust!Address)
            GRDTranx.TextMatrix(i, 4) = IIf(IsNull(rstCust!TELNO), "", rstCust!TELNO)
            GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstCust!KGST), "", rstCust!KGST)
            GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstCust!Area), "", rstCust!Area)
            GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstCust!OPEN_DB), "0.00", Format(rstCust!OPEN_DB, "0.00"))
            GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstCust!AGENT_NAME), "", rstCust!AGENT_NAME)
            i = i + 1
            rstCust.MoveNext
        Loop
    Else
        GRDTranx.Cols = 8
        If optCategory.value = True Then
            rstCust.Open "SELECT * From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        Else
            rstCust.Open "SELECT * From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        End If
        Do Until rstCust.EOF
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(i, 0) = i
            GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!ACT_CODE), "", rstCust!ACT_CODE)
            GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!ACT_NAME), "", rstCust!ACT_NAME)
            GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstCust!Address), "", rstCust!Address)
            GRDTranx.TextMatrix(i, 4) = IIf(IsNull(rstCust!TELNO), "", rstCust!TELNO)
            GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstCust!KGST), "", rstCust!KGST)
            GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstCust!Area), "", rstCust!Area)
            GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstCust!OPEN_DB), "0.00", Format(rstCust!OPEN_DB, "0.00"))
            i = i + 1
            rstCust.MoveNext
        Loop
    End If
    rstCust.Close
    Set rstCust = Nothing
    On Error Resume Next
    GRDTranx.SetFocus
    CMDPRINT.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Function

Private Sub CmbAgent_Click()
    OptAgent.value = True
    CMDPRINT.Enabled = False
End Sub

Private Sub cmbarea_GotFocus()
    optCategory.value = True
    CMDPRINT.Enabled = False
End Sub

Private Sub CmdCashCust_Click()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "NAME"
    GRDTranx.TextMatrix(0, 2) = "PHONE"
    
        
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 4000
    GRDTranx.ColWidth(2) = 2000
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    
    Dim rstTRANX, rstCust As ADODB.Recordset
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    i = 1
    On Error GoTo eRRHAND
    
    Set rstCust = New ADODB.Recordset
    GRDTranx.Cols = 3
    rstCust.Open "SELECT DISTINCT PHONE, BILL_NAME From TRXMAST WHERE PHONE <> '' AND (ACT_CODE = '130000' OR ACT_CODE = '130001') ORDER BY BILL_NAME", db, adOpenForwardOnly
    Do Until rstCust.EOF
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT * FROM TRXMAST WHERE PHONE = '" & rstCust!PHONE & "'", db, adOpenForwardOnly
'        If Not (rstTRANX.EOF Or rstTRANX.BOF) Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(i, 0) = i
            GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!BILL_NAME), "", rstCust!BILL_NAME)
            'GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!BILL_ADDRESS), "", rstCust!BILL_ADDRESS)
            GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!PHONE), "", rstCust!PHONE)
            i = i + 1
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
        
        rstCust.MoveNext
    Loop
    rstCust.Close
    Set rstCust = Nothing
    On Error Resume Next
    GRDTranx.SetFocus
    CMDPRINT.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub CmdDisplay_Click()
    Call Fillgrid
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "CUSTOMER SUPPLIER LIST"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Export") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
    On Error GoTo eRRHAND
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
        oWS.Range("A1", "H1").Merge
        oWS.Range("A1", "H1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "H2").Merge
        oWS.Range("A2", "H2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    
    
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

'    Range("C2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("C:C").EntireColumn.AutoFit     '-- autofitted column
'
'
'    Range("D2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("D:D").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("E2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("E:E").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("F2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("F:F").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("G2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("G:G").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("H2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("H:H").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("I2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("I:I").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("J2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("J:J").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("K2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("K:K").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("L2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("L:L").EntireColumn.AutoFit     '-- autofitted column

'    oWB.ActiveSheet.Font.Name = "Arial"
'    oApp.ActiveSheet.Font.Name = "Arial"
'    oWB.Font.Size = "11"
'    oWB.Font.Bold = True
    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
    If Optcustomers.value = True Then
        oWS.Range("A" & 2).value = "CUSTOMER LIST"
    Else
        oWS.Range("A" & 2).value = "SUPPLIER LIST"
    End If
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 3).value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 3).value = GRDTranx.TextMatrix(0, 2)
    On Error Resume Next
    oWS.Range("D" & 3).value = GRDTranx.TextMatrix(0, 3)
    oWS.Range("E" & 3).value = GRDTranx.TextMatrix(0, 4)
    oWS.Range("F" & 3).value = GRDTranx.TextMatrix(0, 5)
    oWS.Range("G" & 3).value = GRDTranx.TextMatrix(0, 6)
    oWS.Range("H" & 3).value = GRDTranx.TextMatrix(0, 7)
    
    On Error GoTo eRRHAND
    
    i = 4
    For n = 1 To GRDTranx.Rows - 1
        oWS.Range("A" & i).value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).value = GRDTranx.TextMatrix(n, 2)
        If GRDTranx.Cols > 3 Then
            oWS.Range("D" & i).value = GRDTranx.TextMatrix(n, 3)
            oWS.Range("E" & i).value = GRDTranx.TextMatrix(n, 4)
            oWS.Range("F" & i).value = GRDTranx.TextMatrix(n, 5)
            oWS.Range("G" & i).value = GRDTranx.TextMatrix(n, 6)
            oWS.Range("H" & i).value = GRDTranx.TextMatrix(n, 7)
        End If
        On Error GoTo eRRHAND
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    oWS.Columns("A:Z").EntireColumn.AutoFit
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
eRRHAND:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub CmdPrint_Click()
    Dim i As Integer
    
    On Error GoTo eRRHAND
    If Optcustomers.value = True Then
        ReportNameVar = Rptpath & "RptCustList"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.ACT_CODE} <> '130001'))"
        Set CRXFormulaFields = Report.FormulaFields
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
            If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
                Set oRs = New ADODB.Recordset
                Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
                Report.Database.Tables(i).SetDataSource oRs, 3
                Set oRs = Nothing
            End If
        Next i
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
            If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'CUSTOMER DETAILS'"
        Next
    Else
        ReportNameVar = Rptpath & "RptSuppList"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        '(Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3)
        Report.RecordSelectionFormula = "(Mid({CUSTMAST.ACT_CODE}, 1, 3)='311' AND LENGTH({CUSTMAST.ACT_CODE})>3)"
        Set CRXFormulaFields = Report.FormulaFields
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        Next i
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM ACTMAST ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
        
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
            If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'SUPPLIER DETAILS'"
        Next
    End If
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Command2_Click()
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "NAME"
    GRDTranx.TextMatrix(0, 2) = "PHONE"
    
        
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 4000
    GRDTranx.ColWidth(2) = 2000
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    
    Dim rstTRANX, rstCust As ADODB.Recordset
    Dim i As Long
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    i = 1
    On Error GoTo eRRHAND
    
    Set rstCust = New ADODB.Recordset
    GRDTranx.Cols = 3
    rstCust.Open "SELECT DISTINCT PHONE, BILL_NAME From TRXMAST WHERE ACT_CODE = '130000' OR ACT_CODE = '130001' ORDER BY BILL_NAME", db, adOpenForwardOnly
    Do Until rstCust.EOF
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT * FROM TRXMAST WHERE PHONE = '" & rstCust!PHONE & "'", db, adOpenForwardOnly
'        If Not (rstTRANX.EOF Or rstTRANX.BOF) Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(i, 0) = i
            GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!BILL_NAME), "", rstCust!BILL_NAME)
            'GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!BILL_ADDRESS), "", rstCust!BILL_ADDRESS)
            GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!PHONE), "", rstCust!PHONE)
            i = i + 1
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
        
        rstCust.MoveNext
    Loop
    rstCust.Close
    Set rstCust = Nothing
    On Error Resume Next
    GRDTranx.SetFocus
    CMDPRINT.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Sub
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description

End Sub

Private Sub Form_Load()
    
    
    Call FillCombos
    CMDPRINT.Enabled = False
    Call Fillgrid
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub GRDTranx_Click()
    TXTsample.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub GRDTranx_Scroll()
    TXTsample.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub OptAll_Click()
    CMDPRINT.Enabled = False
End Sub

Private Sub OptAllCategory_Click()
    CMDPRINT.Enabled = False
End Sub

Private Sub OptBAL_Click()
    CMDPRINT.Enabled = False
End Sub

Private Sub optCategory_Click()
    CMDPRINT.Enabled = False
End Sub

Private Sub OptCrPeriod_Click()
    CMDPRINT.Enabled = False
End Sub

Private Sub Optcustomers_Click()
    Call FillCombos
End Sub

Private Sub OptSupplier_Click()
    Call FillCombos
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdDisplay_Click
            TxtCode.SetFocus
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdDisplay_Click
            TxtName.SetFocus
    End Select
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Sl_no As Long
    Dim rststock As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 2  ' NAME
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!ACT_NAME = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!ACT_NAME = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing

                    End If
                    GRDTranx.SetFocus
                    
                Case 3  ' ADDRESS
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!Address = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!Address = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    End If
                    GRDTranx.SetFocus

                Case 4  ' PHONE
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!TELNO = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!TELNO = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    End If
                    GRDTranx.SetFocus

                Case 5  ' GST
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!KGST = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!KGST = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    End If
                    GRDTranx.SetFocus
    
                Case 6  ' Area
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!Area = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!Area = Trim(TXTsample.Text)
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    End If
                    GRDTranx.SetFocus
                Case 7  ' OP BAL
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Val(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    If Optcustomers.value = True Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from CUSTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!OPEN_DB = Format(Val(TXTsample.Text), "0.00")
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    Else
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * from ACTMAST where ACT_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        Do Until rststock.EOF
                            rststock!OPEN_DB = Format(Val(TXTsample.Text), "0.00")
                            rststock.Update
                            rststock.MoveNext
                        Loop
                        rststock.Close
                        Set rststock = Nothing
                    End If
                    GRDTranx.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 7
             Select Case KeyAscii
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 2, 3, 4, 5, 6
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDTranx.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDTranx.Col
                    Case 2, 3, 4, 5, 6, 7
                        TXTsample.Visible = True
                        TXTsample.Top = GRDTranx.CellTop + 90
                        TXTsample.Left = GRDTranx.CellLeft
                        TXTsample.Width = GRDTranx.CellWidth
                        TXTsample.Height = GRDTranx.CellHeight
                        TXTsample.Text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                        TXTsample.SetFocus
                End Select
            End If
    End Select
End Sub

Private Function FillCombos()
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo eRRHAND
    If Optcustomers.value = True Then
        cmbarea.Clear
        Set RSTCOMPANY = New ADODB.Recordset
        RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        Do Until RSTCOMPANY.EOF
            If Not IsNull(RSTCOMPANY!Area) Then cmbarea.AddItem (RSTCOMPANY!Area)
            RSTCOMPANY.MoveNext
        Loop
        RSTCOMPANY.Close
        Set RSTCOMPANY = Nothing
        
        CmbAgent.Clear
        Set RSTCOMPANY = New ADODB.Recordset
        RSTCOMPANY.Open "Select DISTINCT AGENT_NAME From CUSTMAST ORDER BY AGENT_NAME", db, adOpenStatic, adLockReadOnly
        Do Until RSTCOMPANY.EOF
            If Not IsNull(RSTCOMPANY!AGENT_NAME) Then CmbAgent.AddItem (RSTCOMPANY!AGENT_NAME)
            RSTCOMPANY.MoveNext
        Loop
        RSTCOMPANY.Close
        Set RSTCOMPANY = Nothing
    Else
        cmbarea.Clear
        Set RSTCOMPANY = New ADODB.Recordset
        RSTCOMPANY.Open "Select DISTINCT AREA From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        Do Until RSTCOMPANY.EOF
            If Not IsNull(RSTCOMPANY!Area) Then cmbarea.AddItem (RSTCOMPANY!Area)
            RSTCOMPANY.MoveNext
        Loop
        RSTCOMPANY.Close
        Set RSTCOMPANY = Nothing
        CmbAgent.Clear
    End If
    Exit Function
eRRHAND:
    MsgBox Err.Description, vbOKOnly, "EzBiz"
End Function
