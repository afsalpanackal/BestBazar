VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMCHQRET2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cheque Return"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7095
   Icon            =   "Frmchqret1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   7095
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   4065
      TabIndex        =   10
      Top             =   1890
      Width           =   2955
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
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
         Left            =   75
         TabIndex        =   0
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
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
         Left            =   1590
         TabIndex        =   1
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00E0E0E0&
      Height          =   4860
      Left            =   45
      TabIndex        =   3
      Top             =   -15
      Width           =   7020
      Begin VB.Frame FrmBank 
         Height          =   2190
         Left            =   30
         TabIndex        =   19
         Top             =   2655
         Width           =   6960
         Begin VB.Frame Frame2 
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1065
            Left            =   4620
            TabIndex        =   26
            Top             =   810
            Visible         =   0   'False
            Width           =   2250
            Begin VB.OptionButton OptNEFT 
               Caption         =   "NEFT / RTGS etc"
               Height          =   195
               Left            =   75
               TabIndex        =   29
               Top             =   750
               Width           =   1770
            End
            Begin VB.OptionButton OptUPI 
               Caption         =   "UPI"
               Height          =   195
               Left            =   75
               TabIndex        =   28
               Top             =   495
               Width           =   1485
            End
            Begin VB.OptionButton optChq 
               Caption         =   "Cheque / Draft"
               Height          =   195
               Left            =   75
               TabIndex        =   27
               Top             =   270
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin MSComCtl2.DTPicker DtChqDate 
            Height          =   360
            Left            =   5325
            TabIndex        =   24
            Top             =   165
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   82837505
            CurrentDate     =   41452
         End
         Begin VB.TextBox TxtChqNo 
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
            ForeColor       =   &H00FF00FF&
            Height          =   345
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   21
            Top             =   210
            Width           =   3510
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1215
            Left            =   1080
            TabIndex        =   25
            Top             =   630
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   2143
            _Version        =   393216
            Appearance      =   0
            Style           =   1
            ForeColor       =   255
            Text            =   ""
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   7
            Left            =   105
            TabIndex        =   23
            Top             =   705
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trnx/ Ref No."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   435
            Index           =   6
            Left            =   90
            TabIndex        =   22
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   5
            Left            =   4770
            TabIndex        =   20
            Top             =   210
            Width           =   540
         End
      End
      Begin VB.OptionButton OptBank 
         BackColor       =   &H00E0E0E0&
         Caption         =   "To Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   345
         Left            =   150
         TabIndex        =   18
         Top             =   2235
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.TextBox TXTREFNO 
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   3990
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1395
         Width           =   2910
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   420
         TabIndex        =   11
         Top             =   765
         Width           =   795
      End
      Begin VB.TextBox txtrcptamt 
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   1470
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1425
         Width           =   1770
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5355
         TabIndex        =   15
         Top             =   765
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711935
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label LBLSUPPLIER 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1140
         TabIndex        =   17
         Top             =   195
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   3375
         TabIndex        =   14
         Top             =   1425
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Returned Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   3840
         TabIndex        =   8
         Top             =   765
         Width           =   1545
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   765
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Return Amt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   1425
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   1245
         TabIndex        =   5
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label LBLDATE 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   360
         Left            =   2580
         TabIndex        =   4
         Top             =   765
         Width           =   1200
      End
   End
   Begin VB.Label lblactcode 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblactcode"
      Height          =   315
      Left            =   1065
      TabIndex        =   12
      Top             =   3210
      Width           =   1620
   End
   Begin VB.Label lbltmprcptamt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "rcpt amount"
      Height          =   315
      Left            =   3150
      TabIndex        =   9
      Top             =   3285
      Width           =   1620
   End
End
Attribute VB_Name = "FRMCHQRET2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean

Private Sub cmdcancel_Click()
    'CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Proper Date", vbOKOnly, "Cheque Return"
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Cheque Return"
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    
    If Val(txtrcptamt.text) = 0 Then
        MsgBox "Enter Returned Cheque Amount", vbOKOnly, "Cheque Return"
        txtrcptamt.SetFocus
        Exit Sub
    End If
    
    If CMBDISTI.BoundText = "" Then
        MsgBox "Please Select the Name of Bank", vbOKOnly, "Cheque Return"
        CMBDISTI.SetFocus
        Exit Sub
    End If
    
'    If OptBank.value = True And DateValue(DtChqDate.value) > DateValue(Date) And ChkStatus.value = 1 Then
'        MsgBox "Please check the status of the Cheque", vbOKOnly, "Cheque Return"
'        ChkStatus.SetFocus
'        Exit Sub
'    End If
    
    On Error GoTo ErrHand
    db.BeginTrans
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RD'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    
    Dim TRX_NO As Double
    TRX_NO = 1
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'RD' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        TRX_NO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From BANK_TRX", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "DR"
    RSTTRXFILE!TRX_NO = TRX_NO
    RSTTRXFILE!BILL_TRX_TYPE = "RD"
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
    RSTTRXFILE!BANK_NAME = CMBDISTI.text
    'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
    RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.text)
    RSTTRXFILE!act_code = lblactcode.Caption
    RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
    'RSTTRXFILE!INV_DATE = LBLDATE.Caption
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
    RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    RSTTRXFILE!CHQ_DATE = Format(DtChqDate.Value, "DD/MM/YYYY")
    RSTTRXFILE!BANK_FLAG = "Y"
    RSTTRXFILE!CHECK_FLAG = "N"
    RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.text)
    If optChq.Value = True Then
        RSTTRXFILE!BANK_MODE = "C"
    ElseIf OptUPI.Value = True Then
        RSTTRXFILE!BANK_MODE = "U"
    ElseIf OptNEFT.Value = True Then
        RSTTRXFILE!BANK_MODE = "N"
    Else
        RSTTRXFILE!BANK_MODE = "C"
    End If
    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "RD"
    RSTTRXFILE!CR_NO = Val(txtBillNo.text)
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
    RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!RCPT_AMT = 0
    RSTTRXFILE!act_code = lblactcode.Caption
    RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
    RSTTRXFILE!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
    RSTTRXFILE!INV_AMT = Val(txtrcptamt.text)
    'RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    RSTTRXFILE!BANK_FLAG = "Y"
    RSTTRXFILE!B_TRX_TYPE = "DR"
    RSTTRXFILE!B_TRX_NO = TRX_NO
    RSTTRXFILE!B_BILL_TRX_TYPE = "RD"
    RSTTRXFILE!B_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
    RSTTRXFILE!BANK_NAME = CMBDISTI.text
    RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.text)
    RSTTRXFILE!CHQ_DATE = Format(DtChqDate.Value, "DD/MM/YYYY")
    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
'    Dim BillNO As Long
'    Set rstBILL = New ADODB.Recordset
'    rstBILL.Open "Select MAX(RCPT_NO) From TRNXRCPT WHERE TRX_TYPE = 'RD'", db, adOpenForwardOnly
'    If Not (rstBILL.EOF And rstBILL.BOF) Then
'        BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
'    End If
'    rstBILL.Close
'    Set rstBILL = Nothing
'
'    Dim SEL_AMOUNT As Double
'    SEL_AMOUNT = Val(txtrcptamt.Text)
'    For i = 0 To FRMPaymntreg.grdcount.Rows - 1
'        If Val(FRMPaymntreg.grdcount.TextMatrix(i, 22)) = 0 Then GoTo SKIP
'        If SEL_AMOUNT <= 0 Then GoTo SKIP
'        BillNO = BillNO + 1
'        Set RSTTRXFILE = New ADODB.Recordset
'        RSTTRXFILE.Open "Select * From TRNXRCPT ", db, adOpenStatic, adLockOptimistic, adCmdText
'        RSTTRXFILE.AddNew
'        RSTTRXFILE!TRX_TYPE = "RD"
'        RSTTRXFILE!RCPT_NO = BillNO
'        RSTTRXFILE!INV_NO = Val(FRMPaymntreg.grdcount.TextMatrix(i, 3))
'        RSTTRXFILE!INV_TRX_TYPE = FRMPaymntreg.grdcount.TextMatrix(i, 8)
'        RSTTRXFILE!INV_TRX_YEAR = Val(FRMPaymntreg.grdcount.TextMatrix(i, 18))
'        RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'        If SEL_AMOUNT > Val(FRMPaymntreg.grdcount.TextMatrix(i, 22)) Then
'            RSTTRXFILE!RCPT_AMOUNT = Val(FRMPaymntreg.grdcount.TextMatrix(i, 22))
'        Else
'            RSTTRXFILE!RCPT_AMOUNT = SEL_AMOUNT
'        End If
'        RSTTRXFILE!ACT_CODE = lblactcode.Caption
'        RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
'        RSTTRXFILE!RCPT_ENTRY_DATE = Format(Date, "DD/MM/YYYY")
'        RSTTRXFILE!REF_NO = Trim(TXTREFNO.Text)
'        RSTTRXFILE!INV_DATE = Format(FRMPaymntreg.grdcount.TextMatrix(i, 2), "DD/MM/YYYY")
'        RSTTRXFILE!CR_NO = Val(txtBillNo.Text)
'        RSTTRXFILE!CR_TRX_TYPE = "CR"
'        RSTTRXFILE.Update
'        SEL_AMOUNT = SEL_AMOUNT - Val(FRMPaymntreg.grdcount.TextMatrix(i, 22))
'        RSTTRXFILE.Close
'        Set RSTTRXFILE = Nothing
'SKIP:
'    Next i
    
       
    db.CommitTrans
    MsgBox "Saved Successfully....", vbOKOnly, "PAYMENT"
    Unload Me
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtrcptamt.SetFocus
    End Select
End Sub

Private Sub Form_Activate()
      txtrcptamt.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ErrHand
    
    AGNT_FLAG = True
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RD'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing

    Call fillcombo
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    DtChqDate.Value = Date
    'Width = 8900
    'Height = 4485
    Left = 800
    Top = 1000
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
    FRMRcptReg.Enabled = True
    FRMRcptReg.GRDTranx.SetFocus
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                txtrcptamt.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                txtrcptamt.SetFocus
            End If
    End Select
End Sub

Private Sub txtrcptamt_GotFocus()
    txtrcptamt.SelStart = 0
    txtrcptamt.SelLength = Len(txtrcptamt.text)
End Sub

Private Sub txtrcptamt_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(txtrcptamt.text) = 0 Then
                MsgBox "Enter Returned Cheque Amount", vbOKOnly, "Cheque Return"
                txtrcptamt.SetFocus
                Exit Sub
            End If
            TXTREFNO.SetFocus
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub txtrcptamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREFNO_GotFocus()
    TXTREFNO.SelStart = 0
    TXTREFNO.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTREFNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdSAVE.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Function fillcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "BANK_NAME"
    CMBDISTI.BoundColumn = "BANK_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function
