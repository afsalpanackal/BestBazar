VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMTRIAL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRIAL BALANCE"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMTRIAL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   10470
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   8730
      Left            =   -120
      TabIndex        =   0
      Top             =   -285
      Width           =   10575
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
         Caption         =   "TRIAL BALANCE"
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
         Height          =   915
         Left            =   135
         TabIndex        =   3
         Top             =   285
         Width           =   10395
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
            Height          =   480
            Left            =   7545
            TabIndex        =   13
            Top             =   330
            Width           =   1380
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
            Height          =   480
            Left            =   9000
            TabIndex        =   12
            Top             =   330
            Width           =   1335
         End
         Begin VB.CommandButton CMDREGISTER 
            Caption         =   "&EXPORT"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   204
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   5985
            TabIndex        =   11
            Top             =   330
            Width           =   1515
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0C0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   4
            Top             =   420
            Value           =   -1  'True
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   5
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   92405761
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   6
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   92405761
            CurrentDate     =   40498
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "FROM"
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
            Left            =   1110
            TabIndex        =   10
            Top             =   405
            Width           =   555
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
            TabIndex        =   9
            Top             =   405
            Width           =   285
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6465
            TabIndex        =   8
            Top             =   1965
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   8685
            TabIndex        =   7
            Top             =   1905
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6990
         Left            =   150
         TabIndex        =   1
         Top             =   1200
         Width           =   10380
         _ExtentX        =   18309
         _ExtentY        =   12330
         _Version        =   393216
         Rows            =   1
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
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
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   465
         Left            =   150
         TabIndex        =   2
         Tag             =   "5"
         Top             =   8205
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   820
         Picture         =   "FRMTRIAL.frx":030A
         ForeColor       =   0
         BarPicture      =   "FRMTRIAL.frx":0326
         Max             =   150
         Text            =   "PLEASE WAIT..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
   End
End
Attribute VB_Name = "FRMTRIAL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmDDisplay_Click()
    'Exit Sub
    GRDTranx.FixedRows = 0
    GRDTranx.Rows = 1
    
    On Error GoTo ErrHand
    Dim Rcptamt As Double
    Dim PymntAmt As Double
    Dim opamt As Double
    
    Rcptamt = 0
    PymntAmt = 0
    opamt = 0
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim CASHAMT, DBTAMT, DIFFAMT, DIFFAMTPY As Double
    
    DIFFAMT = 0
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select DISTINCT act_code From CUSTMAST WHERE act_code <> '130000' AND act_code <> '130001' ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
'    Do Until RSTTRXFILE.EOF
'        CASHAMT = 0
'        Set RSTSALEREG = New ADODB.Recordset
'        RSTSALEREG.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
'            CASHAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
'        End If
'        RSTSALEREG.Close
'        Set RSTSALEREG = Nothing
'
'        DBTAMT = 0
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "'  AND RCPT_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE ='RT' AND (BANK_FLAG <> 'Y' OR ISNULL(BANK_FLAG))", db, adOpenStatic, adLockReadOnly
'        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'            DBTAMT = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'
'        If CASHAMT - DBTAMT <> 0 Then
'            DIFFAMT = DIFFAMT + (CASHAMT - DBTAMT)
'        End If
'
''        If CASHAMT - DBTAMT <> 0 Then
''            Set RSTSALEREG = New ADODB.Recordset
''            RSTSALEREG.Open "SELECT * FROM CASHATRXFILE WHERE AMOUNT = " & DIFFAMT & " AND ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockOptimistic, adCmdText
''            If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
''                'TAXSALEAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
''                RSTSALEREG!AMOUNT = 0
''                RSTSALEREG.Update
''            Else
''                MsgBox ""
''            End If
''            RSTSALEREG.Close
''            Set RSTSALEREG = Nothing
''
''            'MsgBox RSTTRXFILE!ACT_CODE & CASHAMT - DBTAMT
''        End If
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    'MsgBox DIFFAMT
    
'    DIFFAMTPY = 0
'    CASHAMT = 0
'    DBTAMT = 0
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select DISTINCT act_code From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
'    Do Until RSTTRXFILE.EOF
'        CASHAMT = 0
'        Set RSTSALEREG = New ADODB.Recordset
'        RSTSALEREG.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'P' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'PY' ", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
'            CASHAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
'        End If
'        RSTSALEREG.Close
'        Set RSTSALEREG = Nothing
'
'        DBTAMT = 0
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT SUM(RCPT_AMOUNT) From CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "'  AND RCPT_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE ='PY' AND (BANK_FLAG <> 'Y' OR ISNULL(BANK_FLAG))", db, adOpenStatic, adLockReadOnly
'        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
'            DBTAMT = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
'        End If
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'        If CASHAMT - DBTAMT <> 0 Then
'            DIFFAMTPY = DIFFAMTPY + (CASHAMT - DBTAMT)
'        End If
'        If CASHAMT - DBTAMT <> 0 Then
''            Set RSTSALEREG = New ADODB.Recordset
''            RSTSALEREG.Open "SELECT * FROM CASHATRXFILE WHERE AMOUNT = " & DIFFAMT & " AND ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockOptimistic, adCmdText
''            If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
''                'TAXSALEAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
''                RSTSALEREG!AMOUNT = 0
''                RSTSALEREG.Update
''            Else
''                MsgBox ""
''            End If
''            RSTSALEREG.Close
''            Set RSTSALEREG = Nothing
''
'            'MsgBox RSTTRXFILE!ACT_CODE & CASHAMT - DBTAMT
'        End If
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    
    
    'TRIAL BALANCE CALCULATION
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        opamt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select SUM(OPEN_DB) from ACTMAST  WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            opamt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Rcptamt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' AND TRX_TYPE = 'RL' AND RCPT_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Rcptamt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        PymntAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' AND TRX_TYPE = 'PL' AND RCPT_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            PymntAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        opamt = opamt + (Rcptamt - PymntAmt)
            
        Rcptamt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' AND TRX_TYPE = 'RL' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Rcptamt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
    '    If Rcptamt <> 0 Then
    '        GRDTranx.Rows = GRDTranx.Rows + 1
    '        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Personal Deposit"
    '        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Rcptamt, ".00")
    '    End If
        
        PymntAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' AND TRX_TYPE = 'PL' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            PymntAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
    '    If PymntAmt <> 0 Then
    '        GRDTranx.Rows = GRDTranx.Rows + 1
    '        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Personal Drawings"
    '        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(PymntAmt, ".00")
    '    End If
        
        opamt = opamt + (Rcptamt - PymntAmt)
        
        If opamt <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            If opamt > 0 Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Capital Account" & IIf(IsNull(RSTTRXFILE!ACT_NAME), "", " (" & RSTTRXFILE!ACT_NAME & ")")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(opamt), ".00")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Personal Drawings" & IIf(IsNull(RSTTRXFILE!ACT_NAME), "", " (" & RSTTRXFILE!ACT_NAME & ")")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(opamt), ".00")
            End If
        End If
        
        '    If opamt <> 0 Then
    '        GRDTranx.Rows = GRDTranx.Rows + 1
    '        If opamt > 0 Then
    '            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Deposit OP" 'IIf(IsNull(RSTTRXFILE!ACT_NAME), "Personal Deposit OP", RSTTRXFILE!ACT_NAME & " Deposit OP")
    '            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(opamt, ".00")
    '        Else
    '            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Drawings OP" 'IIf(IsNull(RSTTRXFILE!ACT_NAME), "Personal Drawings OP", RSTTRXFILE!ACT_NAME & " Drawings OP")
    '            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(opamt, ".00")
    '        End If
    '    End If
    
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim OPVAL, CLOVAL, RCVDVAL, ISSVAL, PAYCASH As Double
    CLOVAL = 0
    OPVAL = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '111001' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select SUM(OPEN_DB) from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='211')And (LENGTH(ACT_CODE)>3) ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OPVAL = OPVAL + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    RCVDVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RCVDVAL = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    ISSVAL = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT)  FROM CASHATRXFILE WHERE VCH_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        ISSVAL = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    OPVAL = Round(OPVAL + (RCVDVAL - ISSVAL), 2)
    
'    If OPVAL <> 0 Then
'        GRDTranx.Rows = GRDTranx.Rows + 1
'        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "OP Cash"
'        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(OPVAL, ".00")
'    End If
    
    PAYCASH = 0
    RCVDVAL = 0

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='S'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RCVDVAL = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND CHECK_FLAG='P'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        PAYCASH = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    CLOVAL = Round(OPVAL + (RCVDVAL - PAYCASH), 2)
    'CLOVAL = (CLOVAL - DIFFAMT) + DIFFAMTPY
    'CLOVAL = Round((RCVDVAL - PAYCASH), 2)

    If CLOVAL <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Cash in Hand"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(CLOVAL, ".00")
    End If
    
'    Dim expense As Double
'    RSTCOMPANY.Open "Select DISTINCT ACT_NAME From TRXEXPENSE ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
'    expense = 0
'    rstTRANX.Open "SELECT * From TRXEXPENSE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'
    Dim Op_Bal, OP_DR, OP_CR, B_TRX As Double
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From BANKCODE ORDER BY BANK_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        Op_Bal = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select OPEN_DB from BANKCODE WHERE BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Op_Bal = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
'        If Op_Bal <> 0 Then
'            GRDTranx.Rows = GRDTranx.Rows + 1
'            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "BANK (Capital)", RSTTRXFILE!BANK_NAME & " (Capital)")
'            If Op_Bal > 0 Then
'                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(Op_Bal), ".00")
'            Else
'                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(Op_Bal), ".00")
'            End If
'        End If
        
        OP_DR = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'DR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_DR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        OP_CR = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'CR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_CR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Op_Bal = Op_Bal + OP_CR - OP_DR
        
        If Op_Bal <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "BANK OP Amt", RSTTRXFILE!BANK_NAME & " OP Amt")
            If Op_Bal > 0 Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(Op_Bal), ".00")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(Op_Bal), ".00")
            End If
        End If
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(TRX_AMOUNT) FROM BANK_TRX WHERE BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'CR'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        PAYCASH = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(TRX_AMOUNT) FROM BANK_TRX WHERE BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'DR'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            PAYCASH = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        CLOVAL = Round(Op_Bal + (RCVDVAL - PAYCASH), 2)
        'CLOVAL = Round(RCVDVAL - PAYCASH, 2)
        If CLOVAL <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "BANK Clo. Amt", RSTTRXFILE!BANK_NAME & " Clo. Amt")
            If CLOVAL > 0 Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(CLOVAL), ".00")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(CLOVAL), ".00")
            End If
        End If
        
        B_TRX = 0       'DEPOSIT
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_FLAG = 'Y' AND BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'CR' and BILL_TRX_TYPE = 'DP'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            B_TRX = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
            
            If B_TRX > 0 Then
                GRDTranx.Rows = GRDTranx.Rows + 1
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "Bank Deposit to Bank", RSTTRXFILE!BANK_NAME & " Bank Deposit to Bank")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(B_TRX), ".00")
            End If
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        B_TRX = 0       'WITHDRAWAL
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_FLAG = 'Y' AND BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'DR' and BILL_TRX_TYPE = 'WD'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            B_TRX = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
            If B_TRX > 0 Then
                GRDTranx.Rows = GRDTranx.Rows + 1
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "Bank Withdrawal from Bank", RSTTRXFILE!BANK_NAME & " Bank Withdrawal from Bank")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(B_TRX), ".00")
            End If
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        B_TRX = 0   'BANK INTEREST
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_FLAG = 'Y' AND BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'CR' and BILL_TRX_TYPE = 'IN'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            B_TRX = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
            If B_TRX > 0 Then
                GRDTranx.Rows = GRDTranx.Rows + 1
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "Bank Interest", RSTTRXFILE!BANK_NAME & " Bank Interest")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Abs(B_TRX), ".00")
            End If
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        B_TRX = 0   'BANK CHARGE
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(TRX_AMOUNT) from BANK_TRX WHERE BANK_FLAG = 'Y' AND BANK_CODE = '" & RSTTRXFILE!BANK_CODE & "' and TRX_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' and TRX_TYPE = 'DR' and BILL_TRX_TYPE = 'BC'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            B_TRX = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
            If B_TRX > 0 Then
                GRDTranx.Rows = GRDTranx.Rows + 1
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!BANK_NAME), "Bank Charges", RSTTRXFILE!BANK_NAME & " Bank Charges")
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Abs(B_TRX), ".00")
            End If
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    Dim EXP, INC As Double
    Dim actmast As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From TRXEXPENSE ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        EXP = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) From TRXEXPENSE WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "'  and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            EXP = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If EXP <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "OFFICE EXPENSE", actmast!ACT_NAME)
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "OFFICE EXPENSE"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(EXP, ".00")
        End If
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From TRXINCOME ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        INC = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(VCH_AMOUNT) From TRXINCOME WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "'  and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='IN'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            INC = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If INC <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "Other Income", actmast!ACT_NAME)
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Other Income"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(INC, ".00")
        End If
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT EXP_CODE From TRXFILE_EXP ORDER BY EXP_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        EXP = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(TRX_TOTAL) From TRXFILE_EXP WHERE EXP_CODE = '" & RSTTRXFILE!EXP_CODE & "'  and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            EXP = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If EXP <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!EXP_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "STAFF EXPENSE", actmast!ACT_NAME)
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "STAFF EXPENSE"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(EXP, ".00")
        End If
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT ACT_CODE From STAFFPYMT ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        EXP = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMOUNT) From STAFFPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "'  and INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EX'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            EXP = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If EXP <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "STAFF EXPENSE", actmast!ACT_NAME)
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "STAFF EXPENSE"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(EXP, ".00")
        End If
        
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim NetAmt As Double
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(VCH_AMOUNT) From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE= 'SR'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Credit Note"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
    End If
    
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(PAY_AMOUNT) From GIFTMAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE= 'GF'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "SAMPLE GOODS"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
        
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "SAMPLE GOODS (EXPENSE)"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
    End If
    
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(PAY_AMOUNT) From DAMAGE_MAST WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE= 'GF'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "DAMAGE GOODS"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
        
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "DAMAGE GOODS (EXPENSE)"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
    End If
    
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(NET_AMOUNT) From PURCAHSERETURN WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE= 'PR' OR TRX_TYPE= 'WP')", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Debit Note"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
    End If
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'CB' AND INV_TRX_TYPE = 'CN'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Credit Note"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
    End If
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE act_code <> '130000' AND act_code <> '130001' and RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE = 'DB' AND INV_TRX_TYPE = 'DN'", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Debit Note"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
    End If
    
    NetAmt = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From TRANSMAST ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        NetAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) From TRANSMAST WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI' OR TRX_TYPE='PW' OR TRX_TYPE='LP')", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If NetAmt <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "Purchase", actmast!ACT_NAME & " Purchase")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Purchase"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    '=========
    NetAmt = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From ASTRXMAST ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        NetAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) From ASTRXMAST WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='AP'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If NetAmt <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "Assets Purchase", actmast!ACT_NAME & " (Assets Purchase)")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Assets Purchase"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    '=========
    NetAmt = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From ASTRXMAST ORDER BY act_code", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        NetAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) From ASTRXMAST WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='EP'", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If NetAmt <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "Assets Purchase (Input Tax Credit)", actmast!ACT_NAME & " (Assets Purchase (Input Tax Credit))")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Assets Purchase(Input Tax Credit)"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(NetAmt, ".00")
        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    '=========
    
    NetAmt = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From TRXMAST WHERE act_code <> '130000' AND act_code <> '130001' ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        NetAmt = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) From TRXMAST WHERE act_code = '" & RSTTRXFILE!ACT_CODE & "' and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='RI' OR TRX_TYPE='WO')", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If NetAmt <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            Set actmast = New ADODB.Recordset
            actmast.Open "Select * From CUSTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (actmast.EOF And actmast.BOF) Then
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(actmast!ACT_NAME), "Sales", actmast!ACT_NAME & " (Sales)")
            Else
                GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Sales"
            End If
            actmast.Close
            Set actmast = Nothing
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    NetAmt = 0
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT SUM(NET_AMOUNT) From TRXMAST WHERE (act_code = '130000' OR act_code = '130001') and VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='HI' OR TRX_TYPE='GI' OR TRX_TYPE='SI' OR TRX_TYPE='RI' OR TRX_TYPE='WO')", db, adOpenStatic, adLockReadOnly
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        NetAmt = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    If NetAmt <> 0 Then
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = "Cash Sales"
        GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(NetAmt, ".00")
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        Op_Bal = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select OPEN_DB from ACTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Op_Bal = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
'        If Op_Bal <> 0 Then
'            GRDTranx.Rows = GRDTranx.Rows + 1
'            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "Supplier (OP)", RSTTRXFILE!ACT_NAME & " (OP)")
'            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Op_Bal, ".00")
'        End If
        
        OP_DR = 0
        Set rstTRANX = New ADODB.Recordset
        'rstTRANX.Open "Select SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CR') and INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' and TRX_TYPE = 'DR' ", db, adOpenStatic, adLockReadOnly, adCmdText
        rstTRANX.Open "Select SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='CR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_DR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        OP_CR = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(RCPT_AMOUNT) from CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE= 'WP') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_CR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Op_Bal = Op_Bal + (OP_DR - OP_CR)
        
        If Op_Bal <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "Sundry Creditors OP. Amt", RSTTRXFILE!ACT_NAME & "Sundry Creditors OP. Amt")
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(Op_Bal, ".00")
        End If
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='CR' and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        PAYCASH = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMOUNT) from CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and (TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE= 'WP') and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            PAYCASH = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        CLOVAL = Round(Op_Bal + (RCVDVAL - PAYCASH), 2)
        'CLOVAL = Round(RCVDVAL - PAYCASH, 2)
        If CLOVAL <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "", RSTTRXFILE!ACT_NAME & " (Sundry Creditors)")
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(CLOVAL, ".00")
        End If
        RSTTRXFILE!YTD_DB = CLOVAL
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From CUSTMAST WHERE act_code <> '130000' AND act_code <> '130001' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        Op_Bal = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "select OPEN_DB from CUSTMAST WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Op_Bal = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
'        If Op_Bal <> 0 Then
'            GRDTranx.Rows = GRDTranx.Rows + 1
'            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "Debtors OP (Capital)", RSTTRXFILE!ACT_NAME & " OP. (Capital)")
'            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Op_Bal, ".00")
'        End If
        
        OP_DR = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(INV_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='DR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_DR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(RCPT_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='DB' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_DR = OP_DR + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        OP_CR = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "Select SUM(RCPT_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            OP_CR = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Op_Bal = Op_Bal + OP_DR - OP_CR
        
        If Op_Bal <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "Debtors OP Amt", RSTTRXFILE!ACT_NAME & " (Sundry Debtors) OP. Amt")
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(Op_Bal, ".00")
        End If
        
        RCVDVAL = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(INV_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='DR' and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and TRX_TYPE ='DB' and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            RCVDVAL = RCVDVAL + IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        PAYCASH = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) from DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' and (TRX_TYPE ='CB' OR TRX_TYPE ='RT' OR TRX_TYPE ='SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER' OR TRX_TYPE = 'PY') and INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND INV_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            PAYCASH = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        CLOVAL = Round(Op_Bal + (RCVDVAL - PAYCASH), 2)
        'CLOVAL = Round(RCVDVAL - PAYCASH, 2)
        If CLOVAL <> 0 Then
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 0) = IIf(IsNull(RSTTRXFILE!ACT_NAME), "", RSTTRXFILE!ACT_NAME & " (Sundry Debtors)")
            GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(CLOVAL, ".00")
        End If
        RSTTRXFILE!YTD_DB = CLOVAL
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    RCVDVAL = 0
    PAYCASH = 0
    Dim i As Long
    For i = 1 To GRDTranx.Rows - 1
        RCVDVAL = RCVDVAL + Val(GRDTranx.TextMatrix(i, 1))
        PAYCASH = PAYCASH + Val(GRDTranx.TextMatrix(i, 2))
    Next i
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.TextMatrix(GRDTranx.Rows - 1, 1) = Format(RCVDVAL, ".00")
    GRDTranx.TextMatrix(GRDTranx.Rows - 1, 2) = Format(PAYCASH, ".00")
    
    If GRDTranx.Rows > 1 Then GRDTranx.FixedRows = 1
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CMDREGISTER_Click()
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
    On Error GoTo ErrHand
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
        oWS.Range("A1", "J1").Merge
        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "J2").Merge
        oWS.Range("A2", "J2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    oWS.Range("I:I").ColumnWidth = 12
    oWS.Range("J:J").ColumnWidth = 12
    oWS.Range("K:K").ColumnWidth = 12
    oWS.Range("L:L").ColumnWidth = 12
    oWS.Range("M:M").ColumnWidth = 12
    oWS.Range("N:N").ColumnWidth = 12
    oWS.Range("O:O").ColumnWidth = 12
    oWS.Range("P:P").ColumnWidth = 12
    oWS.Range("Q:Q").ColumnWidth = 12
    oWS.Range("R:R").ColumnWidth = 12
    oWS.Range("S:S").ColumnWidth = 12
    oWS.Range("T:T").ColumnWidth = 12
    oWS.Range("U:U").ColumnWidth = 12
    oWS.Range("V:V").ColumnWidth = 12
    
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
    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).text
    oWS.Range("A" & 2).Value = "LEDGER ABSTRACT " & DTFROM.Value & " TO " & DTTO.Value
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GRDTranx.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GRDTranx.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GRDTranx.TextMatrix(0, 2)
    On Error GoTo ErrHand
    
    i = 4
    For n = 1 To GRDTranx.Rows - 1
        oWS.Range("A" & i).Value = GRDTranx.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GRDTranx.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GRDTranx.TextMatrix(n, 2)
        On Error GoTo ErrHand
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
    
    'oWS.Range("D" & i + 1).FormulaR1C1 = "=SUM(RC-10:RC-1)"
SKIP:
    oApp.Visible = True
    
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox err.Description
End Sub

Private Sub Command1_Click()
    Exit Sub
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim CASHAMT, DBTAMT, DIFFAMT, DIFFAMTPY As Double
    
    DIFFAMT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From CUSTMAST WHERE act_code <> '130000' AND act_code <> '130001' ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        CASHAMT = 0
        Set RSTSALEREG = New ADODB.Recordset
        RSTSALEREG.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
            CASHAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
        End If
        RSTSALEREG.Close
        Set RSTSALEREG = Nothing
        
        DBTAMT = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMT) From DBTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "'  AND RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE ='RT' AND (BANK_FLAG <> 'Y' OR ISNULL(BANK_FLAG))", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            DBTAMT = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If CASHAMT - DBTAMT <> 0 Then
            DIFFAMT = DIFFAMT + (CASHAMT - DBTAMT)
        End If
        
'        If CASHAMT - DBTAMT <> 0 Then
'            Set RSTSALEREG = New ADODB.Recordset
'            RSTSALEREG.Open "SELECT * FROM CASHATRXFILE WHERE AMOUNT = " & DIFFAMT & " AND ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockOptimistic, adCmdText
'            If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
'                'TAXSALEAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
'                RSTSALEREG!AMOUNT = 0
'                RSTSALEREG.Update
'            Else
'                MsgBox ""
'            End If
'            RSTSALEREG.Close
'            Set RSTSALEREG = Nothing
'
'            'MsgBox RSTTRXFILE!ACT_CODE & CASHAMT - DBTAMT
'        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    MsgBox DIFFAMT
    
    DIFFAMTPY = 0
    CASHAMT = 0
    DBTAMT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select DISTINCT act_code From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        CASHAMT = 0
        Set RSTSALEREG = New ADODB.Recordset
        RSTSALEREG.Open "SELECT SUM(AMOUNT) FROM CASHATRXFILE WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'P' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND INV_TYPE = 'PY' ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
            CASHAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
        End If
        RSTSALEREG.Close
        Set RSTSALEREG = Nothing
        
        DBTAMT = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(RCPT_AMOUNT) From CRDTPYMT WHERE ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "'  AND RCPT_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND RCPT_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE ='PY' AND (BANK_FLAG <> 'Y' OR ISNULL(BANK_FLAG))", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            DBTAMT = IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0))
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        If CASHAMT - DBTAMT <> 0 Then
            DIFFAMTPY = DIFFAMTPY + (CASHAMT - DBTAMT)
        End If
        If CASHAMT - DBTAMT <> 0 Then
'            Set RSTSALEREG = New ADODB.Recordset
'            RSTSALEREG.Open "SELECT * FROM CASHATRXFILE WHERE AMOUNT = " & DIFFAMT & " AND ACT_CODE = '" & RSTTRXFILE!ACT_CODE & "' AND CHECK_FLAG = 'S' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_TYPE = 'RT' AND (INV_TRX_TYPE = 'WO' OR INV_TRX_TYPE = 'RI' OR INV_TRX_TYPE = 'GI' OR INV_TRX_TYPE = 'SI' OR INV_TRX_TYPE = 'SV' OR INV_TRX_TYPE = 'RT' )", db, adOpenStatic, adLockOptimistic, adCmdText
'            If Not (RSTSALEREG.EOF And RSTSALEREG.BOF) Then
'                'TAXSALEAMT = IIf(IsNull(RSTSALEREG.Fields(0)), 0, RSTSALEREG.Fields(0))
'                RSTSALEREG!AMOUNT = 0
'                RSTSALEREG.Update
'            Else
'                MsgBox ""
'            End If
'            RSTSALEREG.Close
'            Set RSTSALEREG = Nothing
'
            'MsgBox RSTTRXFILE!ACT_CODE & CASHAMT - DBTAMT
        End If
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    MsgBox DIFFAMTPY
    
    MsgBox DIFFAMT - DIFFAMTPY
    Exit Sub
End Sub

Private Sub Form_Load()
    GRDTranx.ColWidth(0) = 5000
    GRDTranx.ColWidth(1) = 1600
    GRDTranx.ColWidth(2) = 1600
    GRDTranx.TextMatrix(0, 0) = "Head"
    GRDTranx.TextMatrix(0, 1) = "Cr"
    GRDTranx.TextMatrix(0, 2) = "Dr"
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = 0
End Sub
