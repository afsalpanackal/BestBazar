VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMContra 
   BackColor       =   &H00D2EDBA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contra Entry"
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   ControlBox      =   0   'False
   Icon            =   "FrmContra.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   7305
   Begin VB.Frame Frame3 
      BackColor       =   &H00ECEBCE&
      Height          =   780
      Left            =   4035
      TabIndex        =   13
      Top             =   4860
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
         Left            =   90
         TabIndex        =   4
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "E&xit"
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
         TabIndex        =   5
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00ECEBCE&
      Height          =   5700
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   7290
      Begin VB.CommandButton CmdBank 
         Caption         =   "&Bank Book"
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
         Left            =   90
         TabIndex        =   28
         Top             =   5100
         Width           =   1365
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E6F5FD&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   1860
         Left            =   90
         TabIndex        =   22
         Top             =   2010
         Width           =   5895
         Begin VB.TextBox TXTDEALER1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   855
            TabIndex        =   23
            Top             =   210
            Width           =   4950
         End
         Begin MSDataListLib.DataList DataList1 
            Height          =   1140
            Left            =   855
            TabIndex        =   24
            Top             =   630
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   2011
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   9
            Left            =   90
            TabIndex        =   25
            Top             =   270
            Width           =   945
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00D6FADF&
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   1845
         Left            =   90
         TabIndex        =   18
         Top             =   150
         Width           =   5895
         Begin VB.TextBox TXTDEALER 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   390
            Left            =   855
            TabIndex        =   19
            Top             =   210
            Width           =   4950
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   1140
            Left            =   855
            TabIndex        =   20
            Top             =   630
            Width           =   4950
            _ExtentX        =   8731
            _ExtentY        =   2011
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   11.25
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   4
            Left            =   90
            TabIndex        =   21
            Top             =   270
            Width           =   945
         End
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
         Left            =   3525
         MaxLength       =   20
         TabIndex        =   3
         Top             =   4515
         Width           =   3465
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
         TabIndex        =   0
         Top             =   4110
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
         Left            =   975
         MaxLength       =   8
         TabIndex        =   2
         Top             =   4515
         Width           =   1770
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5340
         TabIndex        =   1
         Top             =   4110
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
         Left            =   2910
         TabIndex        =   15
         Top             =   4560
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Trnx"
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
         Left            =   3870
         TabIndex        =   11
         Top             =   4140
         Width           =   1395
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
         TabIndex        =   10
         Top             =   4140
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         TabIndex        =   9
         Top             =   4560
         Width           =   1035
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
         TabIndex        =   8
         Top             =   4140
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
         TabIndex        =   7
         Top             =   4110
         Width           =   1215
      End
   End
   Begin VB.Label lbldealer1 
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label flagchange1 
      Height          =   315
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblactcode 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblactcode"
      Height          =   315
      Left            =   1065
      TabIndex        =   14
      Top             =   3210
      Width           =   1620
   End
   Begin VB.Label lbltmprcptamt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "rcpt amount"
      Height          =   315
      Left            =   3150
      TabIndex        =   12
      Top             =   3285
      Width           =   1620
   End
End
Attribute VB_Name = "FRMContra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim BANK_REC As New ADODB.Recordset

Private Sub CmdBank_Click()
    FRMBankBook.Show
    FRMBankBook.SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "Select the Bank from which to be debited", vbOKOnly, "Contra Entry"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    If DataList1.BoundText = "" Then
        MsgBox "Select the Bank to which to be credited", vbOKOnly, "Contra Entry"
        TXTDEALER1.SetFocus
        Exit Sub
    End If
    
    If DataList2.BoundText = DataList1.BoundText Then
        MsgBox "Entry cannot be made to the same bank", vbOKOnly, "Contra Entry"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Proper Date", vbOKOnly, "Contra Entry"
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Contra Entry"
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    
    If Val(txtrcptamt.text) = 0 Then
        MsgBox "Please enter the Amount", vbOKOnly, "Contra Entry"
        txtrcptamt.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    db.BeginTrans
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From BANK_TRX WHERE TRX_NO= (SELECT MAX(TRX_NO) FROM BANK_TRX WHERE BILL_TRX_TYPE = 'CT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' )", db, adOpenStatic, adLockOptimistic, adCmdText
    If RSTTRXFILE.RecordCount = 0 Then
        txtBillNo.text = 1
    Else
        txtBillNo.text = RSTTRXFILE!TRX_NO + 1
    End If
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_NO = Val(txtBillNo.text)
    RSTTRXFILE!TRX_TYPE = "DR"
    RSTTRXFILE!BILL_TRX_TYPE = "CT"
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!BANK_CODE = DataList2.BoundText
    RSTTRXFILE!BANK_NAME = DataList2.text
    RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.text)
    RSTTRXFILE!ACT_CODE = DataList1.BoundText
    RSTTRXFILE!ACT_NAME = DataList1.text
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
    RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    RSTTRXFILE!BANK_FLAG = "Y"
    RSTTRXFILE!check_flag = "Y"
    RSTTRXFILE!CHQ_NO = ""
    RSTTRXFILE.Update
'    RSTTRXFILE!C_TRX_TYPE = Null
'    'RSTTRXFILE!C_REC_NO = Null
'    RSTTRXFILE!C_INV_TRX_TYPE = Null
'    RSTTRXFILE!C_INV_TYPE = Null
'    ''RSTTRXFILE!C_INV_NO = Null
'    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_NO = Val(txtBillNo.text)
    RSTTRXFILE!TRX_TYPE = "CR"
    RSTTRXFILE!BILL_TRX_TYPE = "CT"
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!BANK_CODE = DataList1.BoundText
    RSTTRXFILE!BANK_NAME = DataList1.text
    RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.text)
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
    RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    RSTTRXFILE!BANK_FLAG = "Y"
    RSTTRXFILE!check_flag = "Y"
    RSTTRXFILE!CHQ_NO = ""
'    RSTTRXFILE!C_TRX_TYPE = Null
'    'RSTTRXFILE!C_REC_NO = Null
'    RSTTRXFILE!C_INV_TRX_TYPE = Null
'    RSTTRXFILE!C_INV_TYPE = Null
'    ''RSTTRXFILE!C_INV_NO = Null
'    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
    
   
    On Error GoTo ERRHAND
    Set RSTTRXFILE = New ADODB.Recordset
    'rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BANK_CODE='" & lblactcode.Caption & "' AND  TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'WD' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
    RSTTRXFILE.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BILL_TRX_TYPE = 'CT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenForwardOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        txtBillNo.text = IIf(IsNull(RSTTRXFILE.Fields(0)), 1, RSTTRXFILE.Fields(0) + 1)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    frmemain.Visible = True
    txtrcptamt.SetFocus
    Label1(2).Caption = "Withdrawal Amt"
    
    
    
    txtrcptamt.text = ""
    TXTDEALER.text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    
    TXTDEALER1.text = ""
    lbldealer1.Caption = ""
    flagchange1.Caption = ""
    
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    MsgBox "Saved Successfully....", vbOKOnly, "CONTRA ENTRY"
    TXTDEALER.SetFocus
    Exit Sub
ERRHAND:
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
      TXTDEALER.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo ERRHAND

    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BILL_TRX_TYPE = 'CT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
     
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    'Width = 8900
    'Height = 4485
    'Left = 1000
    'Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Top = 0
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_REC.State = 1 Then ACT_REC.Close
    If BANK_REC.State = 1 Then BANK_REC.Close
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
                MsgBox "Enter Payment Amount", vbOKOnly, "Contra Entry"
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
        Case Asc("'")
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
        Case vbKeyEscape
            txtrcptamt.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_REC.State = 1 Then
            ACT_REC.Close
            ACT_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE BANK_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            ACT_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE BANK_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!BANK_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "BANK_NAME"
        DataList2.BoundColumn = "BANK_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
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

Private Sub DataList2_Click()
'    Dim rstCustomer As ADODB.Recordset
'    Dim RSTTRXFILE As ADODB.Recordset
'
'    On Error GoTo eRRhAND
'    Set rstCustomer = New ADODB.Recordset
'    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
'        lbladdress.Caption = DataList2.Text & Chr(13) & Trim(rstCustomer!Address)
'        TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
'        TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
'    Else
'        TxtPhone.Text = ""
'        TXTTIN.Text = ""
'        lbladdress.Caption = ""
'    End If
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
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
            TXTDEALER1.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

'=============
Private Sub TXTDEALER1_Change()
    On Error GoTo ERRHAND
    If flagchange1.Caption <> "1" Then
        If BANK_REC.State = 1 Then
            BANK_REC.Close
            BANK_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE BANK_NAME Like '" & Me.TXTDEALER1.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            BANK_REC.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE BANK_NAME Like '" & Me.TXTDEALER1.text & "%'ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
        If (BANK_REC.EOF And BANK_REC.BOF) Then
            lbldealer1.Caption = ""
        Else
            lbldealer1.Caption = BANK_REC!BANK_NAME
        End If
        Set Me.DataList1.RowSource = BANK_REC
        DataList1.ListField = "BANK_NAME"
        DataList1.BoundColumn = "BANK_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER1_GotFocus()
    TXTDEALER1.SelStart = 0
    TXTDEALER1.SelLength = Len(TXTDEALER1.text)
End Sub

Private Sub TXTDEALER1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
        Case vbKeyEscape
    End Select
End Sub

Private Sub TXTDEALER1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
'    Dim rstCustomer As ADODB.Recordset
'    Dim RSTTRXFILE As ADODB.Recordset
'
'    On Error GoTo eRRhAND
'    Set rstCustomer = New ADODB.Recordset
'    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList1.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
'        lbladdress.Caption = DataList1.Text & Chr(13) & Trim(rstCustomer!Address)
'        TXTTIN.Text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
'        TxtPhone.Text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
'    Else
'        TxtPhone.Text = ""
'        TXTTIN.Text = ""
'        lbladdress.Caption = ""
'    End If
    TXTDEALER1.text = DataList1.text
    lbldealer1.Caption = TXTDEALER1.text
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList1.SetFocus
                Exit Sub
            End If
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            TXTDEALER1.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange1.Caption = 1
    TXTDEALER1.text = lbldealer1.Caption
    DataList1.text = TXTDEALER1.text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
     flagchange1.Caption = ""
End Sub


