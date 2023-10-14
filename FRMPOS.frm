VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMPOS 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4425
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMPOS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      BackColor       =   &H00C0E0FF&
      Height          =   3210
      Left            =   45
      TabIndex        =   0
      Top             =   -30
      Width           =   4335
      Begin VB.TextBox TxtBankAmt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   1950
         MaxLength       =   10
         TabIndex        =   10
         Top             =   765
         Width           =   2325
      End
      Begin VB.TextBox txtrcptamt 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   510
         Left            =   75
         MaxLength       =   8
         TabIndex        =   6
         Top             =   2970
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.OptionButton OPTCREDIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "BANK"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   375
         TabIndex        =   4
         Top             =   750
         Width           =   1470
      End
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00C0E0FF&
         Caption         =   "CASH"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   375
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1230
      End
      Begin VB.CommandButton CMDEXIT 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "&CANCEL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2820
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2565
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "&OK"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1410
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2565
         Width           =   1365
      End
      Begin MSDataListLib.DataCombo CMBDISTI 
         Height          =   1020
         Left            =   90
         TabIndex        =   8
         Top             =   1290
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   1799
         _Version        =   393216
         Appearance      =   0
         Style           =   1
         ForeColor       =   255
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLCASH 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1950
         TabIndex        =   11
         Top             =   240
         Width           =   2325
      End
      Begin VB.Label lblbnkcode 
         Height          =   240
         Left            =   255
         TabIndex        =   9
         Top             =   2295
         Width           =   285
      End
      Begin VB.Label lbltype 
         Height          =   210
         Left            =   3915
         TabIndex        =   7
         Top             =   645
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label LBLBANK 
         Height          =   225
         Left            =   585
         TabIndex        =   5
         Top             =   2910
         Visible         =   0   'False
         Width           =   1650
      End
   End
End
Attribute VB_Name = "FRMPOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean

Private Sub cmdexit_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    
    If OptCredit.Value = True Then
        If Val(TxtBankAmt.text) <= 0 Then
            MsgBox "Receipt amount cannot be less than or equal to zero", vbOKOnly, "Reciept Entry"
            TxtBankAmt.SetFocus
            Exit Sub
        End If
        If Val(TxtBankAmt.text) > Val(txtrcptamt.text) Then
            MsgBox "Receipt amount bigger than Invoice amount", vbOKOnly, "Reciept Entry"
            TxtBankAmt.SetFocus
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    On Error GoTo ErrHand
    creditbill.lblcredit.Caption = "0"
    creditbill.GRDRECEIPT.rows = 1
    creditbill.GRDRECEIPT.TextMatrix(0, 0) = Val(txtrcptamt.text)
    creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
    creditbill.GRDRECEIPT.TextMatrix(1, 0) = "" 'Trim(TXTREFNO.Text)
    If OptCredit.Value = True Then
        If CMBDISTI.BoundText = "" Then
            Me.Enabled = True
            MsgBox "Please select the Bank from the list", vbOKOnly, "Reciept Entry"
            CMBDISTI.SetFocus
            Exit Sub
        End If
'        Dim RSTBANK As ADODB.Recordset
'        Set RSTBANK = New ADODB.Recordset
'        'RSTBANK .Open "select BANK_CODE, BANK_NAME from BANKCODE ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
'        RSTBANK.Open "select * from BANKCODE WHERE BANK_CODE = '1' ", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTBANK.EOF And RSTBANK.BOF) Then
'            lblbank.Caption = IIf(IsNull(RSTBANK!BANK_NAME), "", RSTBANK!BANK_NAME)
'        End If
'        RSTBANK.Close
'        Set RSTBANK = Nothing
'
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(2, 0) = "B"
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(3, 0) = "" 'Trim(TxtChqNo.Text)
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(4, 0) = 1
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(5, 0) = "" 'Format(DtChqDate.value, "DD/MM/YYYY")
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(6, 0) = "Y"
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(7, 0) = lblbank.Caption 'CMBDISTI.Text
        creditbill.lblCBFLAG.Caption = "B"
        creditbill.lblbankamt.Caption = Val(TxtBankAmt.text)
    Else
'        creditbill.GRDRECEIPT.Rows = creditbill.GRDRECEIPT.Rows + 1
'        creditbill.GRDRECEIPT.TextMatrix(2, 0) = "C"
        creditbill.lblCBFLAG.Caption = "C"
        creditbill.lblbankamt.Caption = 0
    End If
    Screen.MousePointer = vbHourglass
    creditbill.Enabled = True
    Screen.MousePointer = vbHourglass
    creditbill.BANKCODE = CMBDISTI.BoundText
    'On Error Resume Next
    If LBLTYPE.Caption = "S" Then
        Call creditbill.AppendSale
    Else
        Call creditbill.Generateprint
    End If
    Screen.MousePointer = vbNormal
    Unload Me
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 98, 66
            OptCredit.Value = True
            Call cmdOK_Click
        Case 99, 67
            optCash.Value = True
            Call cmdOK_Click
    End Select
End Sub

Private Sub Form_Load()
    AGNT_FLAG = True
    Call fillcombo
    cetre Me
    txtrcptamt.text = Val(creditbill.lblnetamount.Caption)
    If creditbill.lblCBFLAG.Caption = "B" And Val(creditbill.lblnetamount.Caption) > 0 Then
        OptCredit.Value = True
        'Label1(7).Visible = True
        CMBDISTI.Visible = True
        CMBDISTI.BoundText = Trim(creditbill.BANKCODE)
        TxtBankAmt.Visible = True
    Else
        optCash.Value = True
        'Label1(7).Visible = False
        CMBDISTI.Visible = False
        TxtBankAmt.Visible = False
        TxtBankAmt.text = ""
    End If
    TxtBankAmt.text = Format(Val(creditbill.lblbankamt.Caption), "0.00")
    Call TxtBankAmt_Change
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_AGNT.State = 1 Then ACT_AGNT.Close
    creditbill.Enabled = True
End Sub

Private Sub optCash_Click()
    'Label1(7).Visible = False
    CMBDISTI.Visible = False
    TxtBankAmt.Visible = False
    TxtBankAmt.text = ""
End Sub

Private Sub OptCredit_Click()
    'Label1(7).Visible = True
    CMBDISTI.Visible = True
    TxtBankAmt.Visible = True
        
    TxtBankAmt.Visible = True
    If CMBDISTI.VisibleCount = 1 Then
        CMBDISTI.BoundText = lblbnkcode.Caption
    End If
    TxtBankAmt.text = Format(Val(txtrcptamt.text), "0.00")
End Sub

Private Function fillcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE POS_FLAG = 'Y' ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE WHERE POS_FLAG = 'Y' ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    If ACT_AGNT.RecordCount = 1 Then
        lblbnkcode.Caption = ACT_AGNT!BANK_CODE
    Else
        lblbnkcode.Caption = ""
    End If
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "BANK_NAME"
    CMBDISTI.BoundColumn = "BANK_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBDISTI.BoundText = "" Then
                MsgBox "Please select the Bank from the list", vbOKOnly, "Expense Entry"
                CMBDISTI.SetFocus
                Exit Sub
            End If
            cmdOK.SetFocus
    End Select
End Sub

Private Sub TxtBankAmt_Change()
    lblcash.Caption = Format(Val(txtrcptamt.text) - Val(TxtBankAmt.text), "0.00")
End Sub

Private Sub TxtBankAmt_GotFocus()
    TxtBankAmt.SelStart = 0
    TxtBankAmt.SelLength = Len(TxtBankAmt.text)
End Sub

Private Sub TxtBankAmt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtBankAmt_LostFocus()
    TxtBankAmt.text = Format(TxtBankAmt.text, "0.00")
End Sub
