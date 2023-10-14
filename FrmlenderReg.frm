VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMLendReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LENDER'S  REGISTER"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9930
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   9930
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8700
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   9990
      Begin VB.CommandButton CmdRcpt 
         Caption         =   "Reciept"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8745
         TabIndex        =   29
         Top             =   1110
         Width           =   1155
      End
      Begin VB.CommandButton CmdPymnt 
         Caption         =   "Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   8745
         TabIndex        =   28
         Top             =   645
         Width           =   1155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   6315
         TabIndex        =   27
         Top             =   1125
         Width           =   1125
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   120
         TabIndex        =   10
         Top             =   7770
         Width           =   9810
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Op. Balance"
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
            Height          =   300
            Index           =   8
            Left            =   1035
            TabIndex        =   19
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label lblOPBal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   990
            TabIndex        =   18
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLPAIDAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5325
            TabIndex        =   16
            Top             =   435
            Width           =   1875
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paid Amt"
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
            Height          =   240
            Index           =   3
            Left            =   5340
            TabIndex        =   15
            Top             =   150
            Width           =   1875
         End
         Begin VB.Label LBLINVAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   3135
            TabIndex        =   14
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd Amt"
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
            Height          =   300
            Index           =   6
            Left            =   3180
            TabIndex        =   13
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label LBLBALAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   7395
            TabIndex        =   12
            Top             =   435
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Amt"
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
            Height          =   315
            Index           =   7
            Left            =   7395
            TabIndex        =   11
            Top             =   150
            Width           =   1815
         End
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   7500
         TabIndex        =   5
         Top             =   1125
         Width           =   1125
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   5100
         TabIndex        =   4
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1515
         Left            =   120
         TabIndex        =   6
         Top             =   45
         Width           =   4935
         Begin VB.TextBox TxtCode 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   60
            TabIndex        =   22
            Top             =   1170
            Visible         =   0   'False
            Width           =   1875
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
            Height          =   360
            Left            =   1455
            TabIndex        =   1
            Top             =   210
            Width           =   3405
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   840
            Left            =   1455
            TabIndex        =   2
            Top             =   585
            Width           =   3405
            _ExtentX        =   6006
            _ExtentY        =   1482
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
         Begin VB.PictureBox rptPRINT 
            Height          =   480
            Left            =   9990
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   17
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "PARTY NAME"
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
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   9
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   7
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   8
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6195
         Left            =   120
         TabIndex        =   3
         Top             =   1575
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   10927
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   6345
         TabIndex        =   23
         Top             =   225
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   61800449
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   8280
         TabIndex        =   24
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   688
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   40498
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
         Index           =   9
         Left            =   7950
         TabIndex        =   26
         Top             =   270
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
         Index           =   10
         Left            =   5130
         TabIndex        =   25
         Top             =   270
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F7 to make Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   5130
         TabIndex        =   21
         Top             =   825
         Width           =   3120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 to make Receipts"
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
         Height          =   300
         Index           =   8
         Left            =   5130
         TabIndex        =   20
         Top             =   600
         Width           =   3120
      End
   End
End
Attribute VB_Name = "FRMLendReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdPymnt_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMPYMNTLENDER.LBLSUPPLIER.Caption = DataList2.text
    FRMPYMNTLENDER.lblactcode.Caption = DataList2.BoundText
    'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMPYMNTLENDER.Show
    FRMPYMNTLENDER.SetFocus
            
End Sub

Private Sub CmdRcpt_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMRcptLenders.LBLSUPPLIER.Caption = DataList2.text
    FRMRcptLenders.lblactcode.Caption = DataList2.BoundText
    'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMRcptLenders.Show
    FRMRcptLenders.SetFocus
End Sub

Private Sub Command1_Click()
    
    Dim OP_Pymnt, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "please Select Party from the List", vbOKOnly, "Lender Register"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Pymnt = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Rcpt = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='RL' OR TRX_TYPE ='PL') and RCPT_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!TRX_TYPE
            Case "PL"
                OP_Pymnt = OP_Pymnt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case "RL"
                OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Op_Bal = OP_Pymnt - OP_Rcpt
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!OPEN_CR = Op_Bal
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim BAL_AMOUNT As Double
    Dim CR_FLAG As Boolean
    CR_FLAG = False
    BAL_AMOUNT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='RL' OR TRX_TYPE ='PL') AND RCPT_DATE <='" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        CR_FLAG = True
        Select Case RSTTRXFILE!TRX_TYPE
            Case "RL"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case "PL"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                'RSTTRXFILE!BAL_AMT = Op_Bal

        End Select
        RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        Op_Bal = 0
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Sleep (300)
    On Error GoTo ErrHand
    ReportNameVar = Rptpath & "RptLendStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    'Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} = 'RL' OR {DBTPYMT.TRX_TYPE} = 'PL')) AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
    Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} = 'RL' OR {DBTPYMT.TRX_TYPE} = 'PL') AND {DBTPYMT.RCPT_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.RCPT_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    'Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Report.DiscardSavedData
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "TRNX DATE"
    GRDTranx.TextMatrix(0, 3) = "TRNX NO"
    GRDTranx.TextMatrix(0, 4) = "RCPT AMT"
    GRDTranx.TextMatrix(0, 5) = "PAID AMT"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "CR NO"
    
    GRDTranx.ColWidth(0) = 1200
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 1700
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 1600
    GRDTranx.ColWidth(7) = 0
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 1500
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.MNUPYMNT.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
'            FRMEMAIN.Enabled = True
'            FRMEBILL.Visible = False
'            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()

End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim E_TABLE As String
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyF6
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Me.Enabled = False
            FRMRcptLenders.LBLSUPPLIER.Caption = DataList2.text
            FRMRcptLenders.lblactcode.Caption = DataList2.BoundText
            'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMRcptLenders.Show
            FRMRcptLenders.SetFocus
        Case vbKeyF7
            Me.Enabled = False
            FRMPYMNTLENDER.LBLSUPPLIER.Caption = DataList2.text
            FRMPYMNTLENDER.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMPYMNTLENDER.Show
            FRMPYMNTLENDER.SetFocus
    End Select
End Sub

Private Sub GRDTranx_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHand
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            If GRDTranx.rows = 1 Then Exit Sub
            If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                    Select Case GRDTranx.TextMatrix(GRDTranx.Row, 0)
                        Case "Payment"
                            db.BeginTrans
                            db.Execute "delete From DBTPYMT WHERE TRX_TYPE='PL' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & ""
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'PL' AND INV_TRX_TYPE = 'PL' AND REC_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 8)) & " "
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'ML' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 9)) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 10) & "'  "
                            db.CommitTrans
                        Case "Receipt"
                            db.BeginTrans
                            db.Execute "delete From DBTPYMT WHERE TRX_TYPE='RL' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & ""
                            db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = 'DR' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'RL' AND INV_TRX_TYPE = 'RL' AND REC_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 8)) & " "
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'ML' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 9)) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 10) & "'  "
                            db.CommitTrans
                    End Select
                    Call Fillgrid
                Else
                    GRDTranx.SetFocus
                End If
            End If
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    TxtCode.text = DataList2.BoundText
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
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

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Long
    
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') AND INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ORDER BY INV_DATE DESC", db, adOpenForwardOnly
'    Do Until rstTRANX.EOF
'        Select Case rstTRANX!TRX_TYPE
'            Case "RL"
'                GRDTranx.TextMatrix(i, 0) = "Receipt"
'                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
'                GRDTranx.CellForeColor = vbRed
'            Case "PL"
'                GRDTranx.TextMatrix(i, 0) = "Payment"
'                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
'                GRDTranx.CellForeColor = vbBlue
'        End Select
'        rstTRANX.MoveNext
'    Loop
'    rstTRANX.Close
'    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!RCPT_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        Select Case rstTRANX!TRX_TYPE
            Case "RL"
                GRDTranx.TextMatrix(i, 0) = "Receipt"
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbRed
            Case "PL"
                GRDTranx.TextMatrix(i, 0) = "Payment"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
        End Select
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!C_REC_NO), "", rstTRANX!C_REC_NO)
        
        GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
        GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        GRDTranx.Row = i
        GRDTranx.Col = 0
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLBALAMT.Caption = Format((Val(lblOPBal.Caption) + Val(LBLINVAMT.Caption)) - Val(LBLPAIDAMT.Caption), "0.00")
    
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function


Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub
