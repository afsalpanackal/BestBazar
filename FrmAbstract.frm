VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMABSTRACT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LEDGER ABSTRACT"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6450
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   6450
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
      Height          =   3675
      Left            =   15
      TabIndex        =   0
      Top             =   -105
      Width           =   6435
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
         Left            =   2640
         TabIndex        =   2
         Top             =   2820
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
         Left            =   4215
         TabIndex        =   1
         Top             =   2835
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1860
         TabIndex        =   3
         Top             =   1305
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   123076609
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   4035
         TabIndex        =   4
         Top             =   1320
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   123076609
         CurrentDate     =   40498
      End
      Begin VB.Label LBLEXPENSES 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Paid Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   5
         Left            =   1755
         TabIndex        =   16
         Top             =   1950
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblpaidcash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   1710
         TabIndex        =   15
         Top             =   2190
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.Label LBLEXPENSES 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Rcvd cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   255
         Index           =   3
         Left            =   3420
         TabIndex        =   14
         Top             =   1950
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label lblrcvdcash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   3390
         TabIndex        =   13
         Top             =   2190
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblcloscash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   345
         Left            =   5115
         TabIndex        =   12
         Top             =   2295
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label LBLEXPENSES 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Closing Cash"
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
         Height          =   255
         Index           =   4
         Left            =   5145
         TabIndex        =   11
         Top             =   2070
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label lblopcash 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   345
         Left            =   105
         TabIndex        =   10
         Top             =   2190
         Visible         =   0   'False
         Width           =   1560
      End
      Begin VB.Label LBLEXPENSES 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Opening Cash"
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
         Height          =   255
         Index           =   2
         Left            =   -15
         TabIndex        =   9
         Top             =   1935
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   8685
         TabIndex        =   8
         Top             =   285
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   6465
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   1380
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
         TabIndex        =   5
         Top             =   1380
         Width           =   1440
      End
   End
End
Attribute VB_Name = "FRMABSTRACT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDREGISTER_Click()
        
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast, rstCustomer As ADODB.Recordset
    Dim BAL_AMOUNT As Double
    Dim CR_FLAG As Boolean
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    lblopcash.Caption = "0.00"
    lblcloscash.Caption = "0.00"
    
    Dim OPVAL, CLOVAL, RCVDVAL, ISSVAL As Double
    CLOVAL = 0
    OPVAL = 0
    
    Screen.MousePointer = vbHourglass
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '111001' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTo.Value, "MM,DD,YYYY")
    
    RCVDVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        RCVDVAL = RCVDVAL + RSTTRXFILE!AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    ISSVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        ISSVAL = ISSVAL + RSTTRXFILE!AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    lblopcash.Caption = Round(OPVAL + (RCVDVAL - ISSVAL), 2)

    lblpaidcash.Caption = 0
    lblrcvdcash.Caption = 0

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM CASHATRXFILE WHERE VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!check_flag
            Case "S"
                lblrcvdcash.Caption = Val(lblrcvdcash.Caption) + RSTTRXFILE!AMOUNT
            Case "P"
                lblpaidcash.Caption = Val(lblpaidcash.Caption) + RSTTRXFILE!AMOUNT
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    lblcloscash.Caption = Round(Val(lblopcash.Caption) + (Val(lblrcvdcash.Caption) - Val(lblpaidcash.Caption)), 2)
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='SR' OR TRX_TYPE ='RW' OR TRX_TYPE ='DR') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
        RSTTRXFILE.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE!OPEN_CR = Op_Bal
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        CR_FLAG = False
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='RT' OR TRX_TYPE ='DR' OR TRX_TYPE ='SR' OR TRX_TYPE ='RW') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
        
        db.Execute "DELETE FROM DBTPYMT WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "' AND TRX_TYPE ='AA'"
        If CR_FLAG = False Then
            Dim MAXNO As Double
            MAXNO = 1
            Set RstCustmast = New ADODB.Recordset
            RstCustmast.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'AA'", db, adOpenForwardOnly
            If Not (RstCustmast.EOF And RstCustmast.BOF) Then
                MAXNO = IIf(IsNull(RstCustmast.Fields(0)), 1, RstCustmast.Fields(0) + 1)
            End If
            RstCustmast.Close
            Set RstCustmast = Nothing
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * FROM DBTPYMT WHERE ACT_CODE = '" & rstCustomer!ACT_CODE & "' AND TRX_TYPE ='AA'", db, adOpenStatic, adLockOptimistic, adCmdText
            If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                RSTTRXFILE.AddNew
                RSTTRXFILE!TRX_TYPE = "AA"
                RSTTRXFILE!CR_NO = MAXNO
            End If
            RSTTRXFILE!INV_TRX_TYPE = ""
    '        RSTTRXFILE!RCPT_DATE = Null
    '        RSTTRXFILE!RCPT_AMT = Null
            RSTTRXFILE!ACT_CODE = rstCustomer!ACT_CODE
            RSTTRXFILE!ACT_NAME = rstCustomer!ACT_NAME
            RSTTRXFILE!INV_DATE = Format(DTFROM.Value, "DD/MM/YYYY")
            RSTTRXFILE!REF_NO = ""
    '        RSTTRXFILE!INV_AMT = Null
    '        RSTTRXFILE!INV_NO = Null
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    '        RSTTRXFILE!C_TRX_TYPE = Null
    '        'RSTTRXFILE!C_REC_NO = Null
    '        RSTTRXFILE!C_INV_TRX_TYPE = Null
    '        RSTTRXFILE!C_INV_TYPE = Null
    '        ''RSTTRXFILE!C_INV_NO = Null
            RSTTRXFILE!BANK_FLAG = "N"
    '        RSTTRXFILE!B_TRX_TYPE = Null
    '        'RSTTRXFILE!B_TRX_NO = Null
    '        RSTTRXFILE!B_BILL_TRX_TYPE = Null
    '        RSTTRXFILE!B_TRX_YEAR = Null
    '        RSTTRXFILE!BANK_CODE = Null
        
            RSTTRXFILE.Update
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
        End If
        rstCustomer.MoveNext
    Loop
    
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        OP_Sale = IIf(IsNull(rstCustomer!OPEN_DB), 0, rstCustomer!OPEN_DB)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
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
            
        rstCustomer!OPEN_CR = Op_Bal
        
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
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
        
        rstCustomer.MoveNext
    Loop
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    Dim OP_DR, OP_CR As Double

    Dim rstTRANX, rstTRANX2 As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_DR = 0
    OP_CR = 0
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "select * from BANKCODE  ORDER BY BANK_CODE ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstTRANX2.EOF
        Op_Bal = IIf(IsNull(rstTRANX2!OPEN_DB), 0, Format(rstTRANX2!OPEN_DB, "0.00"))
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select * from BANK_TRX WHERE BANK_CODE = '" & rstTRANX2!BANK_CODE & "' and TRX_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until rstTRANX.EOF
            Select Case rstTRANX!TRX_TYPE
                Case "DR"
                    OP_DR = OP_DR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
                Case "CR"
                    OP_CR = OP_CR + IIf(IsNull(rstTRANX!TRX_AMOUNT), 0, rstTRANX!TRX_AMOUNT)
            End Select
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        Op_Bal = Op_Bal + OP_CR - OP_DR
        
        rstTRANX2!OPEN_CR = Op_Bal
        rstTRANX2.Update
        
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        '< '" & Format(DTFROM.value, "yyyy/mm/dd") & "'
        RSTTRXFILE.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & rstTRANX2!BANK_CODE & "' AND TRX_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY BNK_SL_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            'RSTTRXFILE!BAL_AMT = Op_Bal
            Select Case RSTTRXFILE!TRX_TYPE
                Case "CR"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                Case Else
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!TRX_AMOUNT), 0, RSTTRXFILE!TRX_AMOUNT)
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
            Op_Bal = 0
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    
    ReportNameVar = Rptpath & "RptAbstract"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'LEDGER ABSTRACT FOR THE PERIOD ' & cHR(13) &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    
    'LEDGER DEBTORS
    Report.OpenSubreport("RptCustStatmnt.rpt").RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='AA' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='DR' OR {DBTPYMT.TRX_TYPE} ='RW' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} = 'EP' OR {DBTPYMT.TRX_TYPE} = 'VR' OR {DBTPYMT.TRX_TYPE} = 'ER') AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    For i = 1 To Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RptCustStatmnt.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RptCustStatmnt.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RptCustStatmnt.rpt").DiscardSavedData
    Report.OpenSubreport("RptCustStatmnt.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RptCustStatmnt.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'LEDGER (SUNDRY DEBTORS)'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
    Next
    
    'LEDGER CREDTORS
    Report.OpenSubreport("RptSupStatmnt.rpt").RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    For i = 1 To Report.OpenSubreport("RptSupStatmnt.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RptSupStatmnt.rpt").Database.Tables(i).SetLogOnInfo strConnection
    Next i
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM CRDTPYMT ")
        Report.OpenSubreport("RptSupStatmnt.rpt").Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ACTMAST ")
        Report.OpenSubreport("RptSupStatmnt.rpt").Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.OpenSubreport("RptSupStatmnt.rpt").DiscardSavedData
    Report.OpenSubreport("RptSupStatmnt.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RptSupStatmnt.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'LEDGER (SUNDRY CREDTORS)'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
    Next
    
    'CASH BOOK
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTo.Value, "MM,DD,YYYY")
    Report.OpenSubreport("RptCashBook.rpt").RecordSelectionFormula = "({CASHATRXFILE.VCH_DATE}<=# " & Format(DTTo.Value, "MM,DD,YYYY") & " #) AND ({CASHATRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RptCashBook.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RptCashBook.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RptCashBook.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RptCashBook.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RptCashBook.rpt").DiscardSavedData
    Report.OpenSubreport("RptCashBook.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RptCashBook.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@opcash}" Then CRXFormulaField.text = "'" & lblopcash.Caption & "' "
        If CRXFormulaField.Name = "{@clocash}" Then CRXFormulaField.text = "'" & lblcloscash.Caption & "' "
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'CASH BOOK'"
    Next
    
    'OFFICE EXPENSE
    Report.OpenSubreport("RPTOfficeExp.rpt").RecordSelectionFormula = "({TRXFILE_EXP.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE_EXP.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTOfficeExp.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTOfficeExp.rpt").Database.Tables(i).SetLogOnInfo strConnection
    Next i
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM TRXEXPMAST ")
        Report.OpenSubreport("RPTOfficeExp.rpt").Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM TRXEXPENSE ")
        Report.OpenSubreport("RPTOfficeExp.rpt").Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.OpenSubreport("RPTOfficeExp.rpt").DiscardSavedData
    Report.OpenSubreport("RPTOfficeExp.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTOfficeExp.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'OFFICE EXPENSES'"
    Next
    
    'STAFF EXPENSE
    Report.OpenSubreport("RPTStaffExp.rpt").RecordSelectionFormula = "({TRXEXP_MAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXEXP_MAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTStaffExp.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTStaffExp.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTStaffExp.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTStaffExp.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTStaffExp.rpt").DiscardSavedData
    Report.OpenSubreport("RPTStaffExp.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTStaffExp.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'STAFF EXPENSES'"
    Next

    'BANK BOOK
    Report.OpenSubreport("RptBANKREPORT.rpt").RecordSelectionFormula = "({BANK_TRX.TRX_TYPE} = 'DR' OR {BANK_TRX.TRX_TYPE} = 'CR') AND ({BANK_TRX.TRX_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {BANK_TRX.TRX_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    For i = 1 To Report.OpenSubreport("RptBANKREPORT.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RptBANKREPORT.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RptBANKREPORT.rpt").DiscardSavedData
    Report.OpenSubreport("RptBANKREPORT.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RptBANKREPORT.rpt").FormulaFields
'    For Each CRXFormulaField In CRXFormulaFields
'        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'BANK BOOK'"
'    Next
    
    
    'Sales Register
    Report.OpenSubreport("RPTSALESREPORT.rpt").RecordSelectionFormula = "((ISNULL({ITEMMAST.UN_BILL}) OR {ITEMMAST.UN_BILL} <> 'Y') AND ({TRXFILE.TRX_TYPE}='GI' or {TRXFILE.TRX_TYPE}='HI')AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTSALESREPORT.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTSALESREPORT.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTSALESREPORT.rpt").DiscardSavedData
    Report.OpenSubreport("RPTSALESREPORT.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTSALESREPORT.rpt").FormulaFields
'    For Each CRXFormulaField In CRXFormulaFields
'        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'BANK BOOK'"
'    Next
    
    'Purchase Register
    Report.OpenSubreport("RPTPURCHASEREPORT.rpt").RecordSelectionFormula = "({RTRXFILE.TRX_TYPE}='PI' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTPURCHASEREPORT.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTPURCHASEREPORT.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTPURCHASEREPORT.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTPURCHASEREPORT.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTPURCHASEREPORT.rpt").DiscardSavedData
    Report.OpenSubreport("RPTPURCHASEREPORT.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTPURCHASEREPORT.rpt").FormulaFields
'    For Each CRXFormulaField In CRXFormulaFields
'        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
'        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'BANK BOOK'"
'    Next
    
    frmreport.Caption = "ABSTRACT REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
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
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11130
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

