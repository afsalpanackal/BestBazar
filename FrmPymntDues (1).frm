VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMPYMNTDUES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT DUES"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmPymntDues.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   6450
   Begin VB.Frame Frmeperiod 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   15
      TabIndex        =   0
      Top             =   -105
      Width           =   6435
      Begin VB.TextBox TXTDEALER 
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
         Height          =   330
         Left            =   1845
         TabIndex        =   5
         Top             =   1275
         Width           =   3720
      End
      Begin VB.OptionButton OPTCUSTOMER 
         BackColor       =   &H00C0C0FF&
         Caption         =   "SUPPLIER"
         Height          =   210
         Left            =   90
         TabIndex        =   4
         Top             =   1320
         Width           =   1320
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00C0C0FF&
         Caption         =   "All"
         Height          =   210
         Left            =   105
         TabIndex        =   3
         Top             =   900
         Value           =   -1  'True
         Width           =   1020
      End
      Begin VB.CommandButton CMDREGISTER 
         Caption         =   "PRINT REPORT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2610
         TabIndex        =   2
         Top             =   2730
         Width           =   1515
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4185
         TabIndex        =   1
         Top             =   2745
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1860
         TabIndex        =   6
         Top             =   330
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   122028033
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   4035
         TabIndex        =   7
         Top             =   345
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   122028033
         CurrentDate     =   40498
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1035
         Left            =   1845
         TabIndex        =   8
         Top             =   1620
         Width           =   3720
         _ExtentX        =   6562
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
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   8685
         TabIndex        =   12
         Top             =   285
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   6465
         TabIndex        =   11
         Top             =   645
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "TO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   270
         Index           =   5
         Left            =   3585
         TabIndex        =   10
         Top             =   405
         Width           =   285
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   270
         Index           =   4
         Left            =   225
         TabIndex        =   9
         Top             =   405
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FRMPYMNTDUES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CMDREGISTER_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    Dim BAL_AMOUNT As Double
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    If OPTCUSTOMER.Value = True Then
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "CR"
                    OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case Else
                    OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        Op_Bal = OP_Sale - OP_Rcpt
            
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!OPEN_CR = Op_Bal
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
        RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            'RSTTRXFILE!BAL_AMT = Op_Bal
            Select Case RSTTRXFILE!TRX_TYPE
                Case "CR"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                Case Else
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
            Op_Bal = 0
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    Else
        Set RstCustmast = New ADODB.Recordset
        RstCustmast.Open "SELECT * FROM ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RstCustmast.EOF
            Op_Bal = 0
            OP_Sale = 0
            OP_Rcpt = 0
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & RstCustmast!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "CR"
                        OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                    Case Else
                        OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            Op_Bal = OP_Sale - OP_Rcpt
                
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE!OPEN_CR = Op_Bal
                RSTTRXFILE.Update
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            BAL_AMOUNT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & RstCustmast!ACT_CODE & "' and ({DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
            RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & RstCustmast!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until RSTTRXFILE.EOF
                'RSTTRXFILE!BAL_AMT = Op_Bal
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "CR"
                        BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                    Case Else
                        BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                        'RSTTRXFILE!BAL_AMT = Op_Bal
        
                End Select
                RSTTRXFILE!BAL_AMT = BAL_AMOUNT
                Op_Bal = 0
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            RstCustmast.MoveNext
        Loop
        
        RstCustmast.Close
        Set RstCustmast = Nothing
    End If
    Screen.MousePointer = vbNormal
    Sleep (300)
    
    ReportNameVar = Rptpath & "RptSupStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTCUSTOMER.Value = True Then
        Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    Else
        Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    End If
    
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM CRDTPYMT ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ACTMAST ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then
            If OPTCUSTOMER.Value = True Then
                CRXFormulaField.text = "'STATEMENT OF ' & '" & UCase(DataList2.text) & "' & CHR(13) &' FOR THE PERIOD FROM ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
            Else
                CRXFormulaField.text = "'STATEMENT FOR THE PERIOD FROM ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
            End If
        End If
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub

Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDREGISTER.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ACT_FLAG = True
    If Month(Date) - 2 > 4 Then
        DTFROM.Value = "01/" & Format(Month(Date) - 2, "00") & "/" & Year(Date)
    Else
        If Year(Date) > Year(MDIMAIN.DTFROM.Value) Then
            DTFROM.Value = "01/12/" & Year(MDIMAIN.DTFROM.Value)
        Else
            DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
        End If
    End If
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Optall.SetFocus
    End Select
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            Optall.SetFocus
    End Select

End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDREGISTER.SetFocus
        Case vbKeyEscape
            Optall.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub


