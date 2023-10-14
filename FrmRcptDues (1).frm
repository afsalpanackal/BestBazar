VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRCPTDUES 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT DUES"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRcptDues.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   7545
   Begin VB.Frame Frmeperiod 
      BackColor       =   &H00FFC0C0&
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
      Height          =   7770
      Left            =   15
      TabIndex        =   0
      Top             =   -105
      Width           =   7530
      Begin VB.OptionButton OptArea 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Area"
         Height          =   210
         Left            =   4575
         TabIndex        =   19
         Top             =   890
         Width           =   1320
      End
      Begin VB.TextBox TXTDEALER4 
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4605
         TabIndex        =   17
         Top             =   1140
         Width           =   2820
      End
      Begin VB.CommandButton CmdPrnAll 
         Caption         =   "&Print All Pending Bills"
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
         Left            =   60
         TabIndex        =   6
         Top             =   2775
         Width           =   1230
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Display &Missing Bills"
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
         Left            =   2700
         TabIndex        =   8
         Top             =   2775
         Width           =   1245
      End
      Begin VB.CommandButton cmdPend 
         Caption         =   "&Print Pending Bills (Till Date)"
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
         Left            =   1335
         TabIndex        =   7
         Top             =   2775
         Width           =   1320
      End
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
         Left            =   840
         TabIndex        =   5
         Top             =   1140
         Width           =   3720
      End
      Begin VB.OptionButton OPTCUSTOMER 
         BackColor       =   &H00FFC0C0&
         Caption         =   "CUSTOMER"
         Height          =   210
         Left            =   840
         TabIndex        =   4
         Top             =   890
         Width           =   1320
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All"
         Height          =   210
         Left            =   150
         TabIndex        =   3
         Top             =   890
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
         Height          =   615
         Left            =   3990
         TabIndex        =   9
         Top             =   2775
         Width           =   1170
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
         Left            =   5205
         TabIndex        =   11
         Top             =   2775
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1860
         TabIndex        =   1
         Top             =   330
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   114098177
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   4035
         TabIndex        =   2
         Top             =   345
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   114098177
         CurrentDate     =   40498
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1230
         Left            =   840
         TabIndex        =   10
         Top             =   1485
         Width           =   3720
         _ExtentX        =   6562
         _ExtentY        =   2170
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
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4260
         Left            =   30
         TabIndex        =   12
         Top             =   3450
         Width           =   7470
         _ExtentX        =   13176
         _ExtentY        =   7514
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
      Begin MSDataListLib.DataList DataList4 
         Height          =   1260
         Left            =   4605
         TabIndex        =   18
         Top             =   1485
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   2223
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
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
      Begin VB.Label flagchange4 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbldealer4 
         Height          =   315
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   8685
         TabIndex        =   16
         Top             =   285
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   6465
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   405
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FRMRCPTDUES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim AREA_REC As New ADODB.Recordset

Private Sub cmdPend_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    db.Execute "Update DBTPYMT set PAID_FLAG = 'N' where isnull(PAID_FLAG) "
    db.Execute "Update DBTPYMT set RCVD_AMOUNT = 0 where isnull(RCVD_AMOUNT) "
    db.Execute "Update DBTPYMT set INV_AMT = 0 where isnull(INV_AMT) "
    
    ReportNameVar = Rptpath & "RptCustPend"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTCUSTOMER.Value = True Then
        Report.RecordSelectionFormula = "((isnull({DBTPYMT.PAID_FLAG}) or {DBTPYMT.PAID_FLAG} <> 'Y') AND {DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and {DBTPYMT.TRX_TYPE} ='DR' and {DBTPYMT.RCVD_AMOUNT} < {DBTPYMT.INV_AMT} AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " #)"
    Else
        Report.RecordSelectionFormula = "((isnull({DBTPYMT.PAID_FLAG}) or {DBTPYMT.PAID_FLAG} <> 'Y') AND {DBTPYMT.TRX_TYPE} ='DR' and {DBTPYMT.RCVD_AMOUNT} < {DBTPYMT.INV_AMT} AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " #)"
    End If
    
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then
            If OPTCUSTOMER.Value = True Then
                CRXFormulaField.text = "'PENDING BILLS OF ' & '" & UCase(DataList2.text) & "' "
            Else
                CRXFormulaField.text = "'PENDING BILLS'"
            End If
        End If
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdPrnAll_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    db.Execute "Update DBTPYMT set PAID_FLAG = 'N' where isnull(PAID_FLAG) "
    db.Execute "Update DBTPYMT set RCVD_AMOUNT = 0 where isnull(RCVD_AMOUNT) "
    db.Execute "Update DBTPYMT set INV_AMT = 0 where isnull(INV_AMT) "
    
    ReportNameVar = Rptpath & "RptCustPend"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTCUSTOMER.Value = True Then
        Report.RecordSelectionFormula = "((isnull({DBTPYMT.PAID_FLAG}) or {DBTPYMT.PAID_FLAG} <> 'Y') AND {DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and {DBTPYMT.TRX_TYPE} ='DR' and {DBTPYMT.RCVD_AMOUNT} < {DBTPYMT.INV_AMT})"
    Else
        Report.RecordSelectionFormula = "((isnull({DBTPYMT.PAID_FLAG}) or {DBTPYMT.PAID_FLAG} <> 'Y') AND {DBTPYMT.TRX_TYPE} ='DR' and {DBTPYMT.RCVD_AMOUNT} < {DBTPYMT.INV_AMT})"
    End If
    
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then
            If OPTCUSTOMER.Value = True Then
                CRXFormulaField.text = "'PENDING BILLS OF ' & '" & UCase(DataList2.text) & "' "
            Else
                CRXFormulaField.text = "'PENDING BILLS'"
            End If
        End If
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CMDREGISTER_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please Select Customer from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    If OptArea.Value = True And DataList4.BoundText = "" Then
        MsgBox "Please Select Area from the List", vbOKOnly, "Statement"
        TXTDEALER4.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Dim BAL_AMOUNT As Double
    Dim CR_FLAG As Boolean
    Dim MAXNO As Double
    Screen.MousePointer = vbHourglass
    If OPTCUSTOMER.Value = True Then
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='RS' OR TRX_TYPE ='DR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DR"
                    OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                Case "DB"
                    OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        Op_Bal = OP_Sale - OP_Rcpt
            
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!OPEN_CR = Op_Bal
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        CR_FLAG = False
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='RS' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            'RSTTRXFILE!BAL_AMT = Op_Bal
            CR_FLAG = True
            Select Case RSTTRXFILE!TRX_TYPE
                Case "DB"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + (IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
            Op_Bal = 0
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        db.Execute "DELETE FROM DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE ='AA'"
        If CR_FLAG = False Then
            MAXNO = 1
            Set RstCustmast = New ADODB.Recordset
            RstCustmast.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'AA'", db, adOpenForwardOnly
            If Not (RstCustmast.EOF And RstCustmast.BOF) Then
                MAXNO = IIf(IsNull(RstCustmast.Fields(0)), 1, RstCustmast.Fields(0) + 1)
            End If
            RstCustmast.Close
            Set RstCustmast = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE ='AA'", db, adOpenStatic, adLockPessimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_TYPE = "AA"
                RSTTRXFILE!CR_NO = MAXNO
            End If
            RSTTRXFILE!INV_TRX_TYPE = ""
            'RSTTRXFILE!RCPT_DATE = Null
            'RSTTRXFILE!RCPT_AMT = Null
            RSTTRXFILE!ACT_CODE = DataList2.BoundText
            RSTTRXFILE!ACT_NAME = DataList2.text
            RSTTRXFILE!INV_DATE = Format(DTFROM.Value, "DD/MM/YYYY")
            RSTTRXFILE!REF_NO = ""
            'RSTTRXFILE!INV_AMT = Null
            'RSTTRXFILE!INV_NO = Null
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            'RSTTRXFILE!C_TRX_TYPE = Null
            'RSTTRXFILE!C_REC_NO = Null
            'RSTTRXFILE!C_INV_TRX_TYPE = Null
            'RSTTRXFILE!C_INV_TYPE = Null
            ''RSTTRXFILE!C_INV_NO = Null
            RSTTRXFILE!BANK_FLAG = "N"
            'RSTTRXFILE!B_TRX_TYPE = Null
            'RSTTRXFILE!B_TRX_NO = Null
            'RSTTRXFILE!B_BILL_TRX_TYPE = Null
            'RSTTRXFILE!B_TRX_YEAR = Null
            'RSTTRXFILE!BANK_CODE = Null
        
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
    Else
        Dim rstcrbal As ADODB.Recordset
        Set RstCustmast = New ADODB.Recordset
        If OptArea.Value = True Then
            RstCustmast.Open "SELECT * FROM CUSTMAST WHERE AREA ='" & DataList4.BoundText & "' AND (ACT_CODE <> '130000' OR ACT_CODE <> '130001') ORDER BY ACT_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
        Else
            RstCustmast.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' OR ACT_CODE <> '130001' ORDER BY ACT_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
        End If
        Do Until RstCustmast.EOF
            Op_Bal = 0
            OP_Sale = 0
            OP_Rcpt = 0
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & RstCustmast!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='RS' OR TRX_TYPE ='DR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "DR"
                        OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                    Case "DB"
                        OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                    Case Else
                        OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                End Select
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            Op_Bal = OP_Sale - OP_Rcpt
                
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE!OPEN_CR = Op_Bal
                RSTTRXFILE.Update
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            CR_FLAG = False
            BAL_AMOUNT = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & RstCustmast!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE='RW' OR TRX_TYPE ='SR' OR TRX_TYPE ='RS' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until RSTTRXFILE.EOF
                'RSTTRXFILE!BAL_AMT = Op_Bal
                CR_FLAG = True
                Select Case RSTTRXFILE!TRX_TYPE
                    Case "DB"
                        BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                    Case Else
                        BAL_AMOUNT = BAL_AMOUNT + Op_Bal + (IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                        'RSTTRXFILE!BAL_AMT = Op_Bal
        
                End Select
                RSTTRXFILE!BAL_AMT = BAL_AMOUNT
                Op_Bal = 0
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            db.Execute "DELETE FROM DBTPYMT WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "' AND TRX_TYPE ='AA'"
            If CR_FLAG = False Then
                MAXNO = 1
                Set rstcrbal = New ADODB.Recordset
                rstcrbal.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'AA'", db, adOpenForwardOnly
                If Not (rstcrbal.EOF And rstcrbal.BOF) Then
                    MAXNO = IIf(IsNull(rstcrbal.Fields(0)), 1, rstcrbal.Fields(0) + 1)
                End If
                rstcrbal.Close
                Set rstcrbal = Nothing
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE ACT_CODE = '" & RstCustmast!ACT_CODE & "' AND TRX_TYPE ='AA'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    RSTTRXFILE.AddNew
                    RSTTRXFILE!TRX_TYPE = "AA"
                    RSTTRXFILE!CR_NO = MAXNO
                End If
                RSTTRXFILE!INV_TRX_TYPE = ""
'                RSTTRXFILE!RCPT_DATE = Null
'                RSTTRXFILE!RCPT_AMT = Null
                RSTTRXFILE!ACT_CODE = RstCustmast!ACT_CODE
                RSTTRXFILE!ACT_NAME = DataList2.text
                RSTTRXFILE!INV_DATE = Format(DTFROM.Value, "DD/MM/YYYY")
                RSTTRXFILE!REF_NO = ""
'                RSTTRXFILE!INV_AMT = Null
'                RSTTRXFILE!INV_NO = Null
                RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
'                RSTTRXFILE!C_TRX_TYPE = Null
                'RSTTRXFILE!C_REC_NO = Null
'                RSTTRXFILE!C_INV_TRX_TYPE = Null
'                RSTTRXFILE!C_INV_TYPE = Null
                ''RSTTRXFILE!C_INV_NO = Null
                RSTTRXFILE!BANK_FLAG = "N"
'                RSTTRXFILE!B_TRX_TYPE = Null
                'RSTTRXFILE!B_TRX_NO = Null
'                RSTTRXFILE!B_BILL_TRX_TYPE = Null
'                RSTTRXFILE!B_TRX_YEAR = Null
'                RSTTRXFILE!BANK_CODE = Null
'
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
            RstCustmast.MoveNext
        Loop
        
        Dim rstCust, rstTRANX As ADODB.Recordset
        Dim OpBal, AC_DB, AC_CR, Total_DB, Total_CR As Double
        
        Set rstCust = New ADODB.Recordset
        rstCust.Open "SELECT * From CUSTMAST", db, adOpenStatic, adLockOptimistic
        Do Until rstCust.EOF
            rstCust!YTD_CR = 0
            rstCust.Update
            rstCust.MoveNext
        Loop
        rstCust.Close
        Set rstCust = Nothing
        
        Set rstCust = New ADODB.Recordset
        rstCust.Open "SELECT * From CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
        Do Until rstCust.EOF
            OpBal = 0
            Total_DB = 0
            Total_CR = 0
            OpBal = IIf(IsNull(rstCust!OPEN_DB), 0, rstCust!OPEN_DB)
            
            Set rstTRANX = New ADODB.Recordset
            rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstCust!ACT_CODE & "' AND (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'RT' OR TRX_TYPE = 'DR' OR TRX_TYPE='RW' OR TRX_TYPE = 'SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') ORDER BY CR_NO ASC, INV_DATE DESC", db, adOpenForwardOnly
            Do Until rstTRANX.EOF
                AC_DB = 0
                AC_CR = 0
                AC_DB = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
                Select Case rstTRANX!check_flag
                    Case "Y"
                        AC_CR = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
                    Case "N"
                        AC_CR = 0 '""IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
                End Select
                Select Case rstTRANX!TRX_TYPE
                    Case "DB"
                        AC_DB = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                    Case "RT"
                        AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                    Case "CB", "SR", "EP", "VC", "ER", "PY", "RW"
                        AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                End Select
                
                Total_DB = Total_DB + AC_DB
                Total_CR = Total_CR + AC_CR
                rstTRANX.MoveNext
            Loop
            rstTRANX.Close
            Set rstTRANX = Nothing
            rstCust!YTD_CR = Round((OpBal + Total_DB) - Total_CR, 2)
            rstCust.Update
            rstCust.MoveNext
        Loop
        rstCust.Close
        Set rstCust = Nothing
        
    End If
    Screen.MousePointer = vbNormal
    Sleep (300)
    
    ReportNameVar = Rptpath & "RptCustStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTCUSTOMER.Value = True Then
        Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "') and ({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RS' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    Else
        If OptArea.Value = True Then
            Report.RecordSelectionFormula = "({custmast.AREA}='" & DataList4.BoundText & "' AND {DBTPYMT.ACT_CODE}<>'130000' AND {DBTPYMT.ACT_CODE}<>'130001') and ({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RS' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        Else
            Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}<>'130000' AND {DBTPYMT.ACT_CODE}<>'130001') and ({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RS' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER' OR {DBTPYMT.TRX_TYPE} = 'PY') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        End If
        'Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} ='DR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
    End If
    
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

Private Sub Command1_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i As Long
    Dim n As Long
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Bill No"
    GRDBILL.TextMatrix(0, 2) = "Date"
    GRDBILL.TextMatrix(0, 3) = "Type"
    GRDBILL.TextMatrix(0, 4) = "Amount"
    GRDBILL.TextMatrix(0, 5) = "Customer"
    GRDBILL.TextMatrix(0, 6) = "Year"
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 1000
    GRDBILL.ColWidth(2) = 1200
    GRDBILL.ColWidth(3) = 0
    GRDBILL.ColWidth(4) = 1200
    GRDBILL.ColWidth(5) = 2500
    GRDBILL.ColWidth(6) = 1200
    
    GRDBILL.ColAlignment(0) = 4
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 4
    GRDBILL.ColAlignment(3) = 4
    GRDBILL.ColAlignment(4) = 4
    GRDBILL.ColAlignment(5) = 1
    GRDBILL.ColAlignment(6) = 4
    
    GRDBILL.FixedRows = 0
    GRDBILL.rows = 1
    i = 1
    n = 1
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Please select the customer from gthe list", vbOKOnly, "EzBiz"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If Optall.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO')  ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
    If rstTRANX.RecordCount > 0 Then
        MDIMAIN.vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        MDIMAIN.vbalProgressBar1.Max = 100
    End If
    Do Until rstTRANX.EOF
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "SELECT * From DBTPYMT WHERE TRX_YEAR='" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE='DR' AND INV_NO = " & rstTRANX!VCH_NO & " AND INV_TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'", db, adOpenForwardOnly
        If (rstBILL.EOF And rstBILL.BOF) Then
            GRDBILL.rows = GRDBILL.rows + 1
            GRDBILL.FixedRows = 1
            GRDBILL.TextMatrix(i, 0) = i
            GRDBILL.TextMatrix(i, 1) = rstTRANX!VCH_NO
            GRDBILL.TextMatrix(i, 2) = IIf(IsNull(rstTRANX!VCH_DATE), "", rstTRANX!VCH_DATE)
            GRDBILL.TextMatrix(i, 3) = rstTRANX!TRX_TYPE
            GRDBILL.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!NET_AMOUNT), "", rstTRANX!NET_AMOUNT)
            GRDBILL.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            GRDBILL.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
            
'            GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!QTY
'            GRDBILL.TextMatrix(i, 5) = Format(RSTTRXFILE!SALES_PRICE * RSTTRXFILE!QTY, "0.00")
'            GRDBILL.TextMatrix(i, 6) = RSTTRXFILE!REF_NO
'            GRDBILL.TextMatrix(i, 7) = IIf(IsNull(RSTTRXFILE!EXP_DATE), "", RSTTRXFILE!EXP_DATE)
            i = i + 1
        End If
        rstBILL.Close
        Set rstBILL = Nothing
        rstTRANX.MoveNext
        n = n + 1
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
        MDIMAIN.vbalProgressBar1.text = n & "out of " & rstTRANX.RecordCount
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    'db.Execute "delete From DBTPYMT WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='DR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TRX_TYPE = 'GI'"
    MDIMAIN.vbalProgressBar1.ShowText = False
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
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
    If ACT_REC.State = 1 Then ACT_REC.Close
    If AREA_REC.State = 1 Then AREA_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_DblClick()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If frmLogin.rs!Level = "5" Then Exit Sub
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
                
    Select Case Trim(GRDBILL.TextMatrix(GRDBILL.Row, 3))
        Case "HI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDBILL.TextMatrix(GRDBILL.Row, 6)) Then Exit Sub
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                If IsFormLoaded(frmsales) <> True Then
                    frmsales.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    frmsales.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    frmsales.Show
                    frmsales.SetFocus
                    Call frmsales.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES1) <> True Then
                    FRMSALES1.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    FRMSALES1.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    FRMSALES1.Show
                    FRMSALES1.SetFocus
                    Call FRMSALES1.txtBillNo_KeyDown(13, 0)
                ElseIf IsFormLoaded(FRMSALES2) <> True Then
                    FRMSALES2.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    FRMSALES2.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                    FRMSALES2.Show
                    FRMSALES2.SetFocus
                    Call FRMSALES2.txtBillNo_KeyDown(13, 0)
                End If
            Else
                If SALESLT_FLAG = "Y" Then
                    If IsFormLoaded(FRMGSTRSM1) <> True Then
                        FRMGSTRSM1.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM1.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM1.Show
                        FRMGSTRSM1.SetFocus
                        Call FRMGSTRSM1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM2) <> True Then
                        FRMGSTRSM2.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM2.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM2.Show
                        FRMGSTRSM2.SetFocus
                        Call FRMGSTRSM2.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTRSM3) <> True Then
                        FRMGSTRSM3.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM3.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTRSM3.Show
                        FRMGSTRSM3.SetFocus
                        Call FRMGSTRSM3.txtBillNo_KeyDown(13, 0)
                    End If
                Else
                    If IsFormLoaded(FRMGSTR) <> True Then
                        FRMGSTR.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR.Show
                        FRMGSTR.SetFocus
                        Call FRMGSTR.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR1) <> True Then
                        FRMGSTR1.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR1.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR1.Show
                        FRMGSTR1.SetFocus
                        Call FRMGSTR1.txtBillNo_KeyDown(13, 0)
                    ElseIf IsFormLoaded(FRMGSTR2) <> True Then
                        FRMGSTR2.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR2.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                        FRMGSTR2.Show
                        FRMGSTR2.SetFocus
                        Call FRMGSTR2.txtBillNo_KeyDown(13, 0)
                    End If
                End If
            End If
        Case "GI"
            If Year(MDIMAIN.DTFROM.Value) <> Val(GRDBILL.TextMatrix(GRDBILL.Row, 6)) Then Exit Sub
            If IsFormLoaded(FRMGST) <> True Then
                FRMGST.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                FRMGST.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                FRMGST.Show
                FRMGST.SetFocus
                Call FRMGST.txtBillNo_KeyDown(13, 0)
            ElseIf IsFormLoaded(FRMGST1) <> True Then
                FRMGST1.txtBillNo.text = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                FRMGST1.LBLBILLNO.Caption = Val(GRDBILL.TextMatrix(GRDBILL.Row, 1))
                FRMGST1.Show
                FRMGST1.SetFocus
                Call FRMGST1.txtBillNo_KeyDown(13, 0)
            End If
        Case "WO"
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description, , "EzBiz"
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
        Case Asc("'")
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
        If ACT_REC.State = 1 Then ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' AND ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        
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
        Case Asc("'")
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

Private Sub TXTDEALER4_Change()
    
    On Error GoTo ErrHand
    If flagchange4.Caption <> "1" Then
        If AREA_REC.State = 1 Then AREA_REC.Close
        AREA_REC.Open "Select DISTINCT AREA From CUSTMAST WHERE AREA Like '" & TXTDEALER4.text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        If (AREA_REC.EOF And AREA_REC.BOF) Then
            lbldealer4.Caption = ""
        Else
            lbldealer4.Caption = AREA_REC!Area
        End If
        Set Me.DataList4.RowSource = AREA_REC
        DataList4.ListField = "AREA"
        DataList4.BoundColumn = "AREA"
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER4_GotFocus()
    OptArea.Value = True
    TXTDEALER4.SelStart = 0
    TXTDEALER4.SelLength = Len(TXTDEALER4.text)
End Sub

Private Sub TXTDEALER4_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList4.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList4.SetFocus
    End Select

End Sub

Private Sub TXTDEALER4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList4_Click()
        
    TXTDEALER4.text = DataList4.text
    lbldealer4.Caption = TXTDEALER4.text

End Sub

Private Sub DataList4_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER4.text) = "" Then Exit Sub
            If IsNull(DataList4.SelectedItem) Then
                MsgBox "Select Area From List", vbOKOnly, "Area List..."
                DataList4.SetFocus
                Exit Sub
            End If
        Case vbKeyEscape
            TXTDEALER4.SetFocus
    End Select
End Sub

Private Sub DataList4_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList4_GotFocus()
    flagchange4.Caption = 1
    TXTDEALER4.text = lbldealer4.Caption
    DataList4.text = TXTDEALER4.text
    Call DataList4_Click
End Sub

Private Sub DataList4_LostFocus()
     flagchange4.Caption = ""
End Sub

