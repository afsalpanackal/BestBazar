VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRMCSBK 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3840
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   7155
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMCSBK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   7155
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrmeMain 
      BackColor       =   &H00FFC0C0&
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
      Height          =   3780
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   7140
      Begin VB.Frame FrmBank 
         Height          =   2445
         Left            =   90
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   6930
         Begin VB.Frame Frame1 
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   4620
            TabIndex        =   15
            Top             =   795
            Width           =   2250
            Begin VB.OptionButton OptNEFT 
               Caption         =   "NEFT / RTGS etc"
               Height          =   195
               Left            =   75
               TabIndex        =   18
               Top             =   750
               Width           =   1770
            End
            Begin VB.OptionButton OptUPI 
               Caption         =   "UPI"
               Height          =   195
               Left            =   75
               TabIndex        =   17
               Top             =   495
               Width           =   1485
            End
            Begin VB.OptionButton optChq 
               Caption         =   "Cheque / Draft"
               Height          =   195
               Left            =   75
               TabIndex        =   16
               Top             =   270
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin VB.TextBox TXTREFNO 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   405
            Left            =   1080
            MaxLength       =   30
            TabIndex        =   13
            Top             =   1920
            Width           =   5790
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
            TabIndex        =   5
            Top             =   210
            Width           =   3510
         End
         Begin VB.CheckBox ChkStatus 
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4800
            TabIndex        =   4
            Top             =   540
            Width           =   1890
         End
         Begin MSComCtl2.DTPicker DtChqDate 
            Height          =   360
            Left            =   5325
            TabIndex        =   6
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
            Format          =   135135233
            CurrentDate     =   41452
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1215
            Left            =   1080
            TabIndex        =   7
            Top             =   660
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
            Caption         =   "Remarks"
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
            Height          =   360
            Index           =   8
            Left            =   75
            TabIndex        =   14
            Top             =   2010
            Width           =   1005
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
            TabIndex        =   10
            Top             =   210
            Width           =   540
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trnx / Ref No."
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
            Left            =   60
            TabIndex        =   9
            Top             =   195
            Width           =   1050
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
            Left            =   45
            TabIndex        =   8
            Top             =   720
            Width           =   645
         End
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "&CANCEL"
         Height          =   585
         Left            =   3900
         TabIndex        =   2
         Top             =   3090
         Width           =   1470
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   585
         Left            =   5655
         TabIndex        =   1
         Top             =   3090
         Width           =   1380
      End
      Begin MSForms.OptionButton OptBank 
         Height          =   420
         Left            =   2010
         TabIndex        =   12
         Top             =   135
         Width           =   1665
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   8388608
         DisplayStyle    =   5
         Size            =   "2937;741"
         Value           =   "0"
         Caption         =   "Bank"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin MSForms.OptionButton Opt_Cash 
         Height          =   420
         Left            =   135
         TabIndex        =   11
         Top             =   135
         Width           =   1365
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   8388608
         DisplayStyle    =   5
         Size            =   "2408;741"
         Value           =   "1"
         Caption         =   "Cash"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   285
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "FRMCSBK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean


Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBDISTI.BoundText = "" Then
                MsgBox "Please select the Bank from the list", vbOKOnly, "Expense Entry"
                CMBDISTI.SetFocus
                Exit Sub
            End If
            TXTREFNO.SetFocus
    End Select
End Sub

Private Sub CmdExit_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ERRHAND
    Me.Enabled = False
    If Opt_Cash.Value = True Then
        creditbill.lblcash.Caption = "Y"
        creditbill.lblremarks.Caption = ""
        creditbill.lblchqno.Caption = ""
        creditbill.lblbankcode.Caption = ""
        creditbill.lblbankname.Caption = ""
        creditbill.lblchqdate.Caption = ""
        creditbill.lblpassflag.Caption = ""
        creditbill.lblmode.Caption = ""
    Else
        If CMBDISTI.BoundText = "" Then
            Me.Enabled = True
            MsgBox "Please select the Bank from the list", vbOKOnly, "Expense Entry"
            CMBDISTI.SetFocus
            Exit Sub
        End If
        If TxtChqNo.text = "" And TXTREFNO.text = "" Then
            Me.Enabled = True
            MsgBox "Please enter either Cheque No / Remarks", vbOKOnly, "Expense Entry"
            TxtChqNo.SetFocus
            Exit Sub
        End If
        Screen.MousePointer = vbHourglass
        creditbill.lblcash.Caption = "N"
        creditbill.lblremarks.Caption = Trim(TXTREFNO.text)
        creditbill.lblchqno.Caption = Trim(TxtChqNo.text)
        creditbill.lblbankcode.Caption = CMBDISTI.BoundText
        creditbill.lblbankname.Caption = CMBDISTI.text
        creditbill.lblchqdate.Caption = Format(DtChqDate.Value, "DD/MM/YYYY")
        If ChkStatus.Value = 0 Then
            creditbill.lblpassflag.Caption = "N"
        Else
            creditbill.lblpassflag.Caption = "Y"
        End If
        If optChq.Value = True Then
            creditbill.lblmode.Caption = "C"
        ElseIf OptUPI.Value = True Then
            creditbill.lblmode.Caption = "U"
        ElseIf OptNEFT.Value = True Then
            creditbill.lblmode.Caption = "N"
        Else
            creditbill.lblmode.Caption = "C"
        End If
    End If
    creditbill.Enabled = True
    'On Error Resume Next
    creditbill.appendpurchase
    Screen.MousePointer = vbNormal
    Unload Me
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    cetre Me
    AGNT_FLAG = True
    Call fillcombo
    If creditbill.lblcash.Caption = "N" Then
        OptBank.Value = True
        TXTREFNO.text = Trim(creditbill.lblremarks.Caption)
        TxtChqNo.text = Trim(creditbill.lblchqno.Caption)
        CMBDISTI.BoundText = Trim(creditbill.lblbankcode.Caption)
        If IsDate(creditbill.lblchqdate.Caption) Then
            DtChqDate.Value = Format(creditbill.lblchqdate.Caption, "DD/MM/YYYY")
        Else
            DtChqDate.Value = Format(creditbill.TXTINVDATE.text, "DD/MM/YYYY")
        End If
        If creditbill.lblpassflag.Caption = "N" Then
            ChkStatus.Value = 0
        Else
            ChkStatus.Value = 1
        End If
        Select Case creditbill.lblmode.Caption
            Case "C"
                optChq.Value = True
            Case "U"
                OptUPI.Value = True
            Case "N"
                OptNEFT.Value = True
            Case Else
                optChq.Value = True
        End Select
    Else
        DtChqDate.Value = Format(creditbill.TXTINVDATE.text, "DD/MM/YYYY")
        Opt_Cash.Value = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
End Sub

Private Sub OptBank_Click()
    FrmBank.Visible = True
End Sub

Private Sub Opt_Cash_Click()
    FrmBank.Visible = False
End Sub

Private Sub TxtChqNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMBDISTI.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_GotFocus()
    TXTREFNO.SelStart = 0
    TXTREFNO.SelLength = Len(TXTREFNO.text)
End Sub

Private Sub TXTREFNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK.SetFocus
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
    On Error GoTo ERRHAND
    
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

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function


