VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMCounterReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COUNTER  REGISTER"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18945
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
   ScaleWidth      =   18945
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8700
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   19050
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
         TabIndex        =   24
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
         TabIndex        =   9
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
            TabIndex        =   18
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
            TabIndex        =   17
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
            TabIndex        =   15
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Sales Amount"
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
            TabIndex        =   12
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
            TabIndex        =   11
            Top             =   435
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Clo. Amount"
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
            TabIndex        =   10
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         Left            =   90
         TabIndex        =   5
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
            TabIndex        =   19
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
            TabIndex        =   16
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Counter Name"
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
            Left            =   45
            TabIndex        =   8
            Top             =   255
            Width           =   1365
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   6
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   7
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   6345
         TabIndex        =   20
         Top             =   225
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   113180673
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   8280
         TabIndex        =   21
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   688
         _Version        =   393216
         Format          =   113180673
         CurrentDate     =   40498
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Height          =   6015
         Left            =   90
         TabIndex        =   25
         Top             =   1455
         Width           =   18885
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
            Height          =   290
            Left            =   9225
            TabIndex        =   27
            Top             =   660
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSFlexGridLib.MSFlexGrid GRDTranx 
            Height          =   6195
            Left            =   15
            TabIndex        =   26
            Top             =   120
            Width           =   18885
            _ExtentX        =   33311
            _ExtentY        =   10927
            _Version        =   393216
            Rows            =   1
            Cols            =   17
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   350
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            BackColorBkg    =   12632256
            FocusRect       =   2
            AllowUserResizing=   3
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   270
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FRMCounterReg"
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

Private Sub CmdExit_Click()
    Unload Me
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
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Pymnt = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM cont_mast WHERE CONT_NAME = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Rcpt = IIf(IsNull(RSTTRXFILE!OP_AMT), 0, RSTTRXFILE!OP_AMT)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from counter_reg Where CONT_NAME ='" & DataList2.BoundText & "' and (TRX_TYPE ='RL' OR TRX_TYPE ='PL') and RCPT_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
    RSTTRXFILE.Open "SELECT * FROM cont_mast WHERE CONT_NAME = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
    RSTTRXFILE.Open "Select * from counter_reg Where CONT_NAME ='" & DataList2.BoundText & "' and (TRX_TYPE ='RL' OR TRX_TYPE ='PL') AND RCPT_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
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
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RptLendStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'rstTRANX.Open "SELECT * From counter_reg WHERE CONT_NAME = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') ORDER BY TRX_DATE DESC", db, adOpenForwardOnly
    'Report.RecordSelectionFormula = "({counter_reg.CONT_NAME}='" & DataList2.BoundText & "' and ({counter_reg.TRX_TYPE} = 'RL' OR {counter_reg.TRX_TYPE} = 'PL')) AND ({counter_reg.TRX_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {counter_reg.TRX_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
    Report.RecordSelectionFormula = "({counter_reg.CONT_NAME}='" & DataList2.BoundText & "' and ({counter_reg.TRX_TYPE} = 'RL' OR {counter_reg.TRX_TYPE} = 'PL') AND {counter_reg.RCPT_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {counter_reg.RCPT_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    'Report.RecordSelectionFormula = "( {TRXMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    Report.DiscardSavedData
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "Sl"
    GRDTranx.TextMatrix(0, 1) = "DATE"
    GRDTranx.TextMatrix(0, 2) = "Op. Amt"
    GRDTranx.TextMatrix(0, 3) = "Op. Cash"
    GRDTranx.TextMatrix(0, 4) = "Op. Bank"
    GRDTranx.TextMatrix(0, 5) = "Bill Amt"
    GRDTranx.TextMatrix(0, 6) = "Bill Cash"
    GRDTranx.TextMatrix(0, 7) = "Bill Bank"
    GRDTranx.TextMatrix(0, 8) = "Expense"
    GRDTranx.TextMatrix(0, 9) = "Exp. Cash"
    GRDTranx.TextMatrix(0, 10) = "Exp. Bank"
    GRDTranx.TextMatrix(0, 11) = "WD Amt"
    GRDTranx.TextMatrix(0, 12) = "WD Cash"
    GRDTranx.TextMatrix(0, 13) = "WD Bank"
    GRDTranx.TextMatrix(0, 14) = "Clo. Amt"
    GRDTranx.TextMatrix(0, 15) = "Clo. Cash"
    GRDTranx.TextMatrix(0, 16) = "Clo. Bank"
    
    GRDTranx.ColWidth(0) = 400
    GRDTranx.ColWidth(1) = 1100
    GRDTranx.ColWidth(2) = 1000
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 1000
    GRDTranx.ColWidth(6) = 1100
    GRDTranx.ColWidth(7) = 1100
    GRDTranx.ColWidth(8) = 1200
    GRDTranx.ColWidth(9) = 1250
    GRDTranx.ColWidth(10) = 1250
    GRDTranx.ColWidth(11) = 1300
    GRDTranx.ColWidth(12) = 1200
    GRDTranx.ColWidth(13) = 1200
    GRDTranx.ColWidth(14) = 1100
    GRDTranx.ColWidth(15) = 1100
    GRDTranx.ColWidth(16) = 1100
    
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    GRDTranx.ColAlignment(9) = 4
    GRDTranx.ColAlignment(10) = 4
    GRDTranx.ColAlignment(11) = 4
    GRDTranx.ColAlignment(12) = 4
    GRDTranx.ColAlignment(13) = 4
    GRDTranx.ColAlignment(14) = 4
    GRDTranx.ColAlignment(15) = 4
    GRDTranx.ColAlignment(16) = 4
    
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 100
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    'DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    
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

Private Sub GRDTranx_Click()
    TXTsample.Visible = False
End Sub

Private Sub GRDTranx_DblClick()
    FrmDenom.Show
    FrmDenom.SetFocus
    FrmDenom.TxtCAmount.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, 15))
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    If GRDTranx.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 11, 12, 13
                    If frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop + 90
                    TXTsample.Left = GRDTranx.CellLeft '+ 50
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    TXTsample.SetFocus
        
            End Select
            
    End Select
End Sub

Private Sub GRDTranx_Scroll()
    TXTsample.Visible = False
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
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            If frmLogin.rs!Level = "5" Then
                ACT_REC.Open "select CONT_NAME from cont_mast WHERE CONT_NAME = '" & system_name & "'  ORDER BY CONT_NAME", db, adOpenForwardOnly
            Else
                ACT_REC.Open "select CONT_NAME from cont_mast WHERE CONT_NAME Like '" & TXTDEALER.text & "%'ORDER BY CONT_NAME", db, adOpenForwardOnly
            End If
            ACT_FLAG = False
        Else
            ACT_REC.Close
            If frmLogin.rs!Level = "5" Then
                ACT_REC.Open "select CONT_NAME from cont_mast WHERE CONT_NAME = '" & system_name & "'  ORDER BY CONT_NAME", db, adOpenForwardOnly
            Else
                ACT_REC.Open "select CONT_NAME from cont_mast WHERE CONT_NAME Like '" & TXTDEALER.text & "%'ORDER BY CONT_NAME", db, adOpenForwardOnly
            End If
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!CONT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "CONT_NAME"
        DataList2.BoundColumn = "CONT_NAME"
    End If
    Exit Sub
ERRHAND:
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
                MsgBox "Select Counter From List", vbOKOnly, "EzBiz"
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
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
'    Set rstTRANX = New ADODB.Recordset
'    rstTRANX.Open "SELECT * From counter_reg WHERE CONT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RL' OR TRX_TYPE = 'PL') AND TRX_DATE < '" & Format(FROMDATE, "yyyy/mm/dd") & "' ORDER BY TRX_DATE DESC", db, adOpenForwardOnly
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
    Dim FROMDATE As Date
    Dim TODATE As Date
    FROMDATE = DTFROM.Value 'Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = DTTo.Value 'Format(DTTO.Value, "MM,DD,YYYY")
    Dim TOT_AMT, TOT_AMT_CASH, TOT_AMT_BANK As Double
    Dim rstdbt2 As ADODB.Recordset
    Do Until FROMDATE > TODATE
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From COUNTER_REG WHERE CONT_NAME = '" & DataList2.BoundText & "' AND TRX_DATE = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ORDER BY TRX_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            TOT_AMT = 0
            TOT_AMT_CASH = 0
            TOT_AMT_BANK = 0
        
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(NET_AMOUNT) from TRXMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(BANK_AMT) from TRXMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT_BANK = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            TOT_AMT_CASH = TOT_AMT - TOT_AMT_BANK
            
            rstTRANX!INV_AMOUNT = TOT_AMT
            rstTRANX!INV_AMOUNT_CASH = TOT_AMT_CASH
            rstTRANX!INV_AMOUNT_BANK = TOT_AMT_BANK
            
            TOT_AMT = 0
            TOT_AMT_CASH = 0
            TOT_AMT_BANK = 0
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(VCH_AMOUNT) from TRXEXPMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(VCH_AMOUNT) from TRXEXPMAST WHERE CASH_FLAG = 'N' AND SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT_BANK = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            TOT_AMT_CASH = TOT_AMT - TOT_AMT_BANK
            
            rstTRANX!EXP_AMOUNT = TOT_AMT
            rstTRANX!EXP_AMOUNT_CASH = TOT_AMT_CASH
            rstTRANX!EXP_AMOUNT_BANK = TOT_AMT_BANK
            
        Else
            rstTRANX.AddNew
            rstTRANX!CONT_NAME = DataList2.BoundText
            rstTRANX!TRX_DATE = Format(FROMDATE, "DD/MM/YYYY")
            
            TOT_AMT = 0
            TOT_AMT_CASH = 0
            TOT_AMT_BANK = 0
        
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(NET_AMOUNT) from TRXMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(BANK_AMT) from TRXMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT_BANK = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            TOT_AMT_CASH = TOT_AMT - TOT_AMT_BANK
            
            rstTRANX!INV_AMOUNT = TOT_AMT
            rstTRANX!INV_AMOUNT_CASH = TOT_AMT_CASH
            rstTRANX!INV_AMOUNT_BANK = TOT_AMT_BANK
            
            TOT_AMT = 0
            TOT_AMT_CASH = 0
            TOT_AMT_BANK = 0
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(VCH_AMOUNT) from TRXEXPMAST WHERE SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            
            Set rstdbt2 = New ADODB.Recordset
            rstdbt2.Open "select SUM(VCH_AMOUNT) from TRXEXPMAST WHERE CASH_FLAG = 'N' AND SYS_NAME = '" & DataList2.BoundText & "' AND VCH_DATE  = '" & Format(FROMDATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (rstdbt2.EOF And rstdbt2.BOF) Then
                TOT_AMT_BANK = IIf(IsNull(rstdbt2.Fields(0)), 0, rstdbt2.Fields(0))
            End If
            rstdbt2.Close
            Set rstdbt2 = Nothing
            TOT_AMT_CASH = TOT_AMT - TOT_AMT_BANK
            
            rstTRANX!EXP_AMOUNT = TOT_AMT
            rstTRANX!EXP_AMOUNT_CASH = TOT_AMT_CASH
            rstTRANX!EXP_AMOUNT_BANK = TOT_AMT_BANK
            
        End If
        rstTRANX.Update
        rstTRANX.Close
        Set rstTRANX = Nothing
        FROMDATE = DateAdd("d", FROMDATE, 1)
    Loop
    
        
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From COUNTER_REG WHERE CONT_NAME = '" & DataList2.BoundText & "' AND TRX_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY TRX_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = Format(rstTRANX!TRX_DATE, "DD/MM/YYYY")
        'GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstTRANX!OP_AMOUNT), "", rstTRANX!OP_AMOUNT)
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!OP_AMOUNT_CASH), "", rstTRANX!OP_AMOUNT_CASH)
        GRDTranx.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!OP_AMOUNT_BANK), "", rstTRANX!OP_AMOUNT_BANK)
        GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!INV_AMOUNT), "", rstTRANX!INV_AMOUNT)
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!INV_AMOUNT_CASH), "", rstTRANX!INV_AMOUNT_CASH)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!INV_AMOUNT_BANK), "", rstTRANX!INV_AMOUNT_BANK)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!EXP_AMOUNT), "", rstTRANX!EXP_AMOUNT)
        GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!EXP_AMOUNT_CASH), "", rstTRANX!EXP_AMOUNT_CASH)
        GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!EXP_AMOUNT_BANK), "", rstTRANX!EXP_AMOUNT_BANK)
        GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!DR_AMOUNT), "", rstTRANX!DR_AMOUNT)
        GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!DR_AMOUNT_CASH), "", rstTRANX!DR_AMOUNT_CASH)
        GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!DR_AMOUNT_BANK), "", rstTRANX!DR_AMOUNT_BANK)
        GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstTRANX!CLO_AMOUNT), "", rstTRANX!CLO_AMOUNT)
        GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstTRANX!CLO_AMOUNT_CASH), "", rstTRANX!CLO_AMOUNT_CASH)
        GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rstTRANX!CLO_AMOUNT_BANK), "", rstTRANX!CLO_AMOUNT_BANK)
        
'        GRDTranx.Row = i
'        GRDTranx.Col = 0
'        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
'        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OP_AMT from CONT_MAST  WHERE CONT_NAME = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OP_AMT), 0, Format(rstTRANX!OP_AMT, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLBALAMT.Caption = Format((Val(lblOPBal.Caption) + Val(LBLINVAMT.Caption)) - Val(LBLPAIDAMT.Caption), "0.00")
    
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
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

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                
                '11,12,13
                Case 11  'WD AMOUNT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * From COUNTER_REG WHERE CONT_NAME = '" & DataList2.BoundText & "' AND TRX_DATE = '" & Format(GRDTranx.TextMatrix(GRDTranx.Row, 1), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!DR_AMOUNT = Val(TXTsample.text)
                        rststock!DR_AMOUNT_CASH = Val(TXTsample.text) - Val(GRDTranx.TextMatrix(GRDTranx.Row, 13))
                        rststock!DR_AMOUNT_BANK = Val(TXTsample.text) - Val(GRDTranx.TextMatrix(GRDTranx.Row, 12))
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        GRDTranx.TextMatrix(GRDTranx.Row, 12) = Format(Val(TXTsample.text) - Val(GRDTranx.TextMatrix(GRDTranx.Row, 13)), "0.00")
                        GRDTranx.TextMatrix(GRDTranx.Row, 13) = Format(Val(TXTsample.text) - Val(GRDTranx.TextMatrix(GRDTranx.Row, 12)), "0.00")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                Case 12  'WD AMOUNT CASH
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * From COUNTER_REG WHERE CONT_NAME = '" & DataList2.BoundText & "' AND TRX_DATE = '" & Format(GRDTranx.TextMatrix(GRDTranx.Row, 1), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!DR_AMOUNT_CASH = Val(TXTsample.text)
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        GRDTranx.TextMatrix(GRDTranx.Row, 11) = Format(Val(GRDTranx.TextMatrix(GRDTranx.Row, 13)) + Val(TXTsample.text), "0.00")
                        
                        rststock!DR_AMOUNT = Val(GRDTranx.TextMatrix(GRDTranx.Row, 11))
                        rststock!DR_AMOUNT_BANK = Val(GRDTranx.TextMatrix(GRDTranx.Row, 13))
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                Case 13  'WD AMOUNT BANK
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * From COUNTER_REG WHERE CONT_NAME = '" & DataList2.BoundText & "' AND TRX_DATE = '" & Format(GRDTranx.TextMatrix(GRDTranx.Row, 1), "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!DR_AMOUNT_BANK = Val(TXTsample.text)
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        GRDTranx.TextMatrix(GRDTranx.Row, 11) = Format(Val(GRDTranx.TextMatrix(GRDTranx.Row, 12)) + Val(TXTsample.text), "0.00")
                        
                        rststock!DR_AMOUNT = Val(GRDTranx.TextMatrix(GRDTranx.Row, 11))
                        rststock!DR_AMOUNT_CASH = Val(GRDTranx.TextMatrix(GRDTranx.Row, 12))
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    
                
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 11, 12, 13
             Select Case KeyAscii
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

