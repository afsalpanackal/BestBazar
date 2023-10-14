VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Frmreminder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT REMINDER"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   16380
   Icon            =   "FrmRemind.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8880
   ScaleWidth      =   16380
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save Receipt Entries"
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
      Left            =   14835
      TabIndex        =   17
      Top             =   7860
      Width           =   1590
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
      Height          =   450
      Left            =   7365
      TabIndex        =   12
      Top             =   7860
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   975
      Left            =   5685
      TabIndex        =   8
      Top             =   -60
      Width           =   10650
      Begin VB.TextBox TxtRef 
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
         Left            =   7830
         TabIndex        =   23
         Top             =   540
         Width           =   2760
      End
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
         TabIndex        =   14
         Top             =   120
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
         TabIndex        =   13
         Top             =   120
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
         TabIndex        =   11
         Top             =   135
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
         TabIndex        =   10
         Top             =   135
         Value           =   -1  'True
         Width           =   1230
      End
      Begin MSComCtl2.DTPicker DTRCPT 
         Height          =   390
         Left            =   8805
         TabIndex        =   15
         Top             =   120
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         Format          =   87359489
         CurrentDate     =   40498
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F7 for 8B Sales"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   300
         Index           =   6
         Left            =   75
         TabIndex        =   28
         Top             =   705
         Width           =   3510
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F2 to enter Receipt Amounts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   5
         Left            =   75
         TabIndex        =   27
         Top             =   480
         Width           =   3510
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref. No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   3
         Left            =   7020
         TabIndex        =   24
         Top             =   585
         Width           =   810
      End
      Begin VB.Label LblInvoice 
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Index           =   1
         Left            =   7485
         TabIndex        =   16
         Top             =   180
         Width           =   1275
      End
      Begin MSForms.ComboBox cmbarea 
         Height          =   360
         Left            =   5685
         TabIndex        =   9
         Top             =   120
         Width           =   1740
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "3069;635"
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
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1005
      Left            =   -30
      TabIndex        =   2
      Top             =   -75
      Width           =   5715
      Begin VB.OptionButton OptCrPeriod 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Over Credit Period"
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
         Left            =   3585
         TabIndex        =   5
         Top             =   195
         Width           =   2070
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "All Customers"
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
         TabIndex        =   4
         Top             =   180
         Width           =   1605
      End
      Begin VB.OptionButton OptBAL 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Oustanding Only"
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
         Left            =   1710
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1875
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
      Left            =   9690
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
      Height          =   450
      Left            =   8505
      TabIndex        =   0
      Top             =   7860
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Height          =   6990
      Left            =   0
      TabIndex        =   20
      Top             =   825
      Width           =   16365
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
         TabIndex        =   22
         Top             =   870
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6855
         Left            =   15
         TabIndex        =   21
         Top             =   90
         Width           =   16350
         _ExtentX        =   28840
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   1
         Cols            =   9
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
   Begin VB.Label LblLastRcpt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   5190
      TabIndex        =   26
      Top             =   7860
      Width           =   2145
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Rcpt Amt"
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
      Height          =   300
      Index           =   4
      Left            =   3780
      TabIndex        =   25
      Top             =   7950
      Width           =   1410
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Amount"
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
      Height          =   300
      Index           =   2
      Left            =   10875
      TabIndex        =   19
      Top             =   7950
      Width           =   1590
   End
   Begin VB.Label LblReceipt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   12465
      TabIndex        =   18
      Top             =   7875
      Width           =   2340
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Height          =   300
      Index           =   0
      Left            =   105
      TabIndex        =   7
      Top             =   7935
      Width           =   1410
   End
   Begin VB.Label lblAMT 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   450
      Left            =   1425
      TabIndex        =   6
      Top             =   7845
      Width           =   2310
   End
End
Attribute VB_Name = "Frmreminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim M_EDIT1, M_EDIT2 As Boolean

Private Function Fillgrid()
    Dim rstTRANX, rstCust As ADODB.Recordset
    Dim OpBal, AC_DB, AC_CR, Total_DB, Total_CR As Double
    Dim DueDays As String
    Dim CR_PERIOD, DUE_DATE, Last_Rcpt_Amt As Long
    Dim i As Integer
    
    If optCategory.value = True And cmbarea.Text = "" Then
        MsgBox "Please select the Place from the List", vbOKOnly, "Receipt Dues"
        On Error Resume Next
        cmbarea.SetFocus
        Exit Function
    End If
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    i = 1
    lblAMT.Caption = ""
    LblLastRcpt.Caption = ""
    LblReceipt.Caption = ""
    On Error GoTo eRRHAND
    
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
    If optCategory.value = True Then
        rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND AREA Like '%" & cmbarea.Text & "%' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
    Else
        rstCust.Open "SELECT * From CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE Like '%" & Trim(TxtCode.Text) & "%' AND ACT_NAME Like '%" & Trim(TxtName.Text) & "%' ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic
    End If
    Do Until rstCust.EOF
        OpBal = 0
        Total_DB = 0
        Total_CR = 0
        DUE_DATE = 0
        DueDays = ""
        Last_Rcpt_Amt = 0
        CR_PERIOD = IIf(IsNull(rstCust!PYMT_PERIOD), 0, rstCust!PYMT_PERIOD)
        OpBal = IIf(IsNull(rstCust!OPEN_DB), 0, rstCust!OPEN_DB)
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & rstCust!ACT_CODE & "' AND (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'RT' OR TRX_TYPE = 'DR' OR TRX_TYPE = 'SR' OR TRX_TYPE = 'EP' OR TRX_TYPE = 'VR' OR TRX_TYPE = 'ER') ORDER BY CR_NO ASC, INV_DATE DESC", db, adOpenForwardOnly
        Do Until rstTRANX.EOF
            AC_DB = 0
            AC_CR = 0
            AC_DB = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
            Select Case rstTRANX!CHECK_FLAG
                Case "Y"
                    AC_CR = IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
                Case "N"
                    AC_CR = 0 '""IIf(IsNull(rstTRANX!INV_AMT), 0, rstTRANX!INV_AMT)
            End Select
            Select Case rstTRANX!TRX_TYPE
                Case "DR"
                    If IsDate(rstTRANX!INV_DATE) Then
                        DUE_DATE = DateDiff("d", rstTRANX!INV_DATE, Date)
                        DueDays = DUE_DATE & " days"
                    End If
                Case "DB"
                    AC_DB = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                Case "RT"
                    AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                    Last_Rcpt_Amt = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
                Case "CB", "SR", "EP", "VC", "ER"
                    AC_CR = IIf(IsNull(rstTRANX!RCPT_AMT), 0, rstTRANX!RCPT_AMT)
            End Select
            
            Total_DB = Total_DB + AC_DB
            Total_CR = Total_CR + AC_CR
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
        
        If Optall.value = False And (OpBal + Total_DB) - Total_CR = 0 Then GoTo SKIP
        If OptCrPeriod.value = True And DUE_DATE < CR_PERIOD Then GoTo SKIP
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = IIf(IsNull(rstCust!ACT_CODE), "", rstCust!ACT_CODE)
        GRDTranx.TextMatrix(i, 2) = IIf(IsNull(rstCust!ACT_NAME), "", rstCust!ACT_NAME)
        GRDTranx.TextMatrix(i, 3) = OpBal + Total_DB
        GRDTranx.TextMatrix(i, 4) = Total_CR
        GRDTranx.TextMatrix(i, 5) = Round((OpBal + Total_DB) - Total_CR, 2)
        rstCust!YTD_CR = Val(GRDTranx.TextMatrix(i, 5))
        rstCust.Update
        GRDTranx.TextMatrix(i, 6) = DueDays
        If Last_Rcpt_Amt > 0 Then GRDTranx.TextMatrix(i, 7) = Last_Rcpt_Amt
        GRDTranx.TextMatrix(i, 8) = ""
        lblAMT.Caption = Format(Val(lblAMT.Caption) + GRDTranx.TextMatrix(i, 5), "0.00")
        LblLastRcpt.Caption = Val(LblLastRcpt.Caption) + Val(GRDTranx.TextMatrix(i, 7))
        i = i + 1
SKIP:
        rstCust.MoveNext
    Loop
    rstCust.Close
    Set rstCust = Nothing
    
    LblLastRcpt.Caption = Format(Round(Val(LblLastRcpt.Caption), 2), "0.00")
    
    DTRCPT.value = Null
    M_EDIT1 = False
    M_EDIT2 = False
    TxtRef.Text = ""
    On Error Resume Next
    GRDTranx.SetFocus
    CMDPRINT.Enabled = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Function

Private Sub cmbarea_GotFocus()
    optCategory.value = True
    CMDPRINT.Enabled = False
End Sub

Private Sub CMDDISPLAY_Click()
    If M_EDIT1 = True And M_EDIT2 = True Then
        If MsgBox("Changes have been made. Do you want to save the changes?", vbYesNo, "Receipt Entries...") = vbNo Then
            Call Fillgrid
            Exit Sub
        Else
            Call CmdSave_Click
            Exit Sub
        End If
    End If
    Call Fillgrid
End Sub

Private Sub CMDEXIT_Click()
    If M_EDIT1 = True And M_EDIT2 = True Then
        If MsgBox("Changes have been made. Do you want to save the changes?", vbYesNo, "Receipt Entries...") = vbYes Then
            Call CmdSave_Click
            If Not (M_EDIT1 = True And M_EDIT2 = True) Then
                Unload Me
            End If
            Exit Sub
        End If
    End If
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim i As Integer
    
    On Error GoTo eRRHAND
    ReportNameVar = MDIMAIN.StatusBar.Panels(7).Text & "EzBiz\RptRecSt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If optCategory.value = True Then
        Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.AREA} startswith '" & cmbarea.Text & "' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(TxtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
    Else
        Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.ACT_CODE} startswith '" & Trim(TxtCode.Text) & "' AND {CUSTMAST.ACT_NAME} startswith '" & Trim(TxtName.Text) & "' AND {CUSTMAST.YTD_CR} <>0))"
    End If
    'Report.RecordSelectionFormula = "(({CUSTMAST.ACT_CODE} <> '130000' AND {CUSTMAST.YTD_CR} <>0))"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", MDIMAIN.StatusBar.Panels(6).Text, "admin", "###DATABASE%%%RET"
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'CUSTOMER DETAILS'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub CmdSave_Click()
    
    Dim RSTTRXFILE, rstBILL As ADODB.Recordset
    Dim Sl_no, Max_Rec, Rec_Nos As Long
    
    If IsNull(DTRCPT.value) Then
        MsgBox "Please select Date of Receipt", vbOKOnly, "Receipt"
        DTRCPT.SetFocus
        Exit Sub
    End If
    
'    If Trim(TxtRef.Text) = "" Then
'        MsgBox "Please enter the Reference No.", vbOKOnly, "Receipt"
'        TxtRef.SetFocus
'        Exit Sub
'    End If
    
    If MsgBox("ARE YOU SURE YOU WANT TO SAVE ALL THE RECEIPT ENTRIES", vbYesNo, "RECEIPT.....") = vbNo Then Exit Sub
    On Error GoTo eRRHAND
    Dim i As Integer
    Dim RECNO, INVNO As Long
    Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    
    Screen.MousePointer = vbHourglass
    Rec_Nos = 0
    LblReceipt.Caption = ""
    For Sl_no = 1 To GRDTranx.Rows - 1
        If Val(GRDTranx.TextMatrix(Sl_no, 8)) = 0 Then GoTo SKIP
        Rec_Nos = Rec_Nos + 1
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(Val(CR_NO)) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            Max_Rec = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
        
        i = 0
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(Val(REC_NO)) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
        'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(creditbill.LBLBILLNO.Caption) & " AND INV_TYPE = 'RT'"
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST.AddNew
        RSTITEMMAST!rec_no = i + 1
        RSTITEMMAST!INV_TYPE = "RT"
        RSTITEMMAST!INV_TRX_TYPE = "RT"
        RSTITEMMAST!INV_NO = Max_Rec
        RSTITEMMAST!TRX_TYPE = "CR"
        RSTITEMMAST!ACT_CODE = GRDTranx.TextMatrix(Sl_no, 1)
        RSTITEMMAST!ACT_NAME = GRDTranx.TextMatrix(Sl_no, 2)
        RSTITEMMAST!AMOUNT = Val(GRDTranx.TextMatrix(Sl_no, 8))
        RSTITEMMAST!VCH_DATE = Format(DTRCPT.value, "DD/MM/YYYY")
        RSTITEMMAST!BILL_TRX_TYPE = "SI"
        RSTITEMMAST!CASH_MODE = "C"
        RSTITEMMAST!CHQ_NO = ""
        RSTITEMMAST!CHQ_DATE = Null
        RSTITEMMAST!BANK = ""
        RSTITEMMAST!CHQ_STATUS = ""
        RSTITEMMAST!CHECK_FLAG = "S"
        RECNO = RSTITEMMAST!rec_no
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "RT"
        RSTTRXFILE!INV_TRX_TYPE = "RI"
        RSTTRXFILE!CR_NO = Max_Rec
        RSTTRXFILE!RCPT_DATE = Format(DTRCPT.value, "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMT = Val(GRDTranx.TextMatrix(Sl_no, 8))
        RSTTRXFILE!ACT_CODE = GRDTranx.TextMatrix(Sl_no, 1)
        RSTTRXFILE!ACT_NAME = GRDTranx.TextMatrix(Sl_no, 2)
        RSTTRXFILE!INV_DATE = Format(DTRCPT.value, "DD/MM/YYYY")
        RSTTRXFILE!REF_NO = Trim(TxtRef.Text)
        RSTTRXFILE!INV_AMT = Null
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTTRXFILE!INV_NO = 0
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!BANK_FLAG = "N"
        RSTTRXFILE!B_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_NO = Null
        RSTTRXFILE!B_BILL_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_YEAR = Null
        RSTTRXFILE!BANK_CODE = Null
        RSTTRXFILE!C_TRX_TYPE = TRXTYPE
        RSTTRXFILE!C_REC_NO = RECNO
        RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
        RSTTRXFILE!C_INV_TYPE = INVTYPE
        RSTTRXFILE!C_INV_NO = INVNO
        
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        LblReceipt.Caption = Val(LblReceipt.Caption) + Val(GRDTranx.TextMatrix(Sl_no, 8))
SKIP:
    Next Sl_no
    LblReceipt.Caption = Format(LblReceipt.Caption, "0.00")
    Screen.MousePointer = vbNormal
    MsgBox Rec_Nos & " Entries Saved", vbOKOnly, "Receipt Entry"
    M_EDIT1 = False
    M_EDIT2 = False
    CMDDISPLAY_Click
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Form_Load()

    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "CODE"
    GRDTranx.TextMatrix(0, 2) = "NAME"
    GRDTranx.TextMatrix(0, 3) = "CREDIT"
    GRDTranx.TextMatrix(0, 4) = "DEBIT"
    GRDTranx.TextMatrix(0, 5) = "BALANCE"
    GRDTranx.TextMatrix(0, 6) = "Last Bill"
    GRDTranx.TextMatrix(0, 7) = "Last Rcpt"
    GRDTranx.TextMatrix(0, 8) = "Rcpt Amt"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 1000
    GRDTranx.ColWidth(2) = 4000
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 1600
    GRDTranx.ColWidth(7) = 1600
    GRDTranx.ColWidth(8) = 1600
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 4
    GRDTranx.ColAlignment(8) = 4
    
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo eRRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!Area) Then cmbarea.AddItem (RSTCOMPANY!Area)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    DTRCPT.value = Format(Date, "DD/MM/YYYY")
    DTRCPT.value = Null
    CMDPRINT.Enabled = False
    Call Fillgrid
    Me.Left = 1000
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub GRDTranx_Click()
    TXTsample.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If GRDTranx.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            For Each F In Forms
                If F.Name = "FRMRECEIPT" Then
                    MsgBox "Please close the Receipt Window", vbOKOnly, "Receipt Entry"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
                If F.Name = "FRMDRCR" Then
                    MsgBox "Please close the Receipt Window", vbOKOnly, "Receipt Entry"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
            Next F
            FRMRcptReg.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
            FRMRcptReg.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMRcptReg.Show
            FRMRcptReg.SetFocus
            
        Case 118
            For Each F In Forms
                If F.Name = "FRMESTIMATE" Then
                    MsgBox "Sales WIndow Already Opened", vbOKOnly, "Sales"
                    GRDTranx.SetFocus
                    Exit Sub
                End If
            Next F
'            FRMRcptReg.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
'            FRMRcptReg.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMESTIMATE.Show
            FRMESTIMATE.SetFocus
            FRMESTIMATE.TXTDEALER.Text = GRDTranx.TextMatrix(GRDTranx.Row, 2)
            FRMESTIMATE.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            FRMESTIMATE.DataList2_Click
            FRMESTIMATE.DataList2.BoundText = GRDTranx.TextMatrix(GRDTranx.Row, 1)
        Case 113
            If frmLogin.rs!Level = "0" Then
                Select Case GRDTranx.Col
                    Case 8
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

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY_Click
            TxtCode.SetFocus
    End Select
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY_Click
            TxtName.SetFocus
    End Select
End Sub

Private Sub TxtName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTsample_Change()
    M_EDIT1 = True
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Sl_no As Long
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 8  ' Rcpt
                    If Val(TXTsample.Text) > Val(GRDTranx.TextMatrix(GRDTranx.Row, 5)) Then
                        MsgBox "Receipt Amount could not be greater than Balance Amount"
                        Exit Sub
                    End If
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Val(TXTsample.Text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    LblReceipt.Caption = ""
                    For Sl_no = 1 To GRDTranx.Rows - 1
                        LblReceipt.Caption = Val(LblReceipt.Caption) + Val(GRDTranx.TextMatrix(Sl_no, 8))
                    Next Sl_no
                    LblReceipt.Caption = Format(LblReceipt.Caption, "0.00")
                    GRDTranx.SetFocus
                    M_EDIT2 = True
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 8
             Select Case KeyAscii
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub
