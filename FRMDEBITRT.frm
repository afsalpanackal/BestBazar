VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRMDEBITRT 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   8280
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMDEBITRT.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      BackColor       =   &H00C0E0FF&
      Height          =   4740
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   8265
      Begin VB.OptionButton OPTCREDIT 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Credit"
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
         Left            =   2595
         TabIndex        =   19
         Top             =   255
         Value           =   -1  'True
         Width           =   1680
      End
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cash"
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
         Left            =   435
         TabIndex        =   18
         Top             =   255
         Width           =   1680
      End
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
         Height          =   3720
         Left            =   45
         TabIndex        =   3
         Top             =   975
         Visible         =   0   'False
         Width           =   8175
         Begin VB.Frame FrmBank 
            Height          =   1980
            Left            =   75
            TabIndex        =   6
            Top             =   1680
            Visible         =   0   'False
            Width           =   6930
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
               Left            =   4860
               TabIndex        =   8
               Top             =   660
               Width           =   1890
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
               TabIndex        =   7
               Top             =   210
               Width           =   3510
            End
            Begin MSComCtl2.DTPicker DtChqDate 
               Height          =   360
               Left            =   5325
               TabIndex        =   9
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
               Format          =   120193025
               CurrentDate     =   41452
            End
            Begin MSDataListLib.DataCombo CMBDISTI 
               Height          =   1215
               Left            =   1080
               TabIndex        =   10
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
               TabIndex        =   13
               Top             =   705
               Width           =   645
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Cheque / Draft No."
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
               TabIndex        =   12
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
               TabIndex        =   11
               Top             =   210
               Width           =   540
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
            Height          =   510
            Left            =   4020
            MaxLength       =   20
            TabIndex        =   5
            Top             =   375
            Width           =   2925
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
            Left            =   1500
            MaxLength       =   8
            TabIndex        =   4
            Top             =   375
            Width           =   1785
         End
         Begin MSForms.OptionButton Opt_Cash 
            Height          =   420
            Left            =   150
            TabIndex        =   17
            Top             =   1065
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
         Begin MSForms.OptionButton OptBank 
            Height          =   420
            Left            =   2010
            TabIndex        =   16
            Top             =   1065
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
            Left            =   3405
            TabIndex        =   15
            Top             =   435
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd Amount"
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
            Left            =   135
            TabIndex        =   14
            Top             =   510
            Width           =   1335
         End
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "&CANCEL"
         Height          =   585
         Left            =   6525
         TabIndex        =   2
         Top             =   255
         Width           =   1470
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   585
         Left            =   4995
         TabIndex        =   1
         Top             =   255
         Width           =   1380
      End
      Begin VB.Label lbltype 
         Height          =   225
         Left            =   135
         TabIndex        =   20
         Top             =   660
         Width           =   330
      End
   End
End
Attribute VB_Name = "FRMDEBITRT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean


Private Sub CmdExit_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    On Error GoTo ERRHAND
    If optCash.Value = True And Val(txtrcptamt.Text) = 0 Then
        MsgBox "Please enter the receipt amount", vbOKOnly, "Sales Bill"
        Exit Sub
    End If
    If optCash.Value = True And OptBank.Value = True And CMBDISTI.BoundText = "" Then
        MsgBox "Please Select the Bank from the list", vbOKOnly, "Sales Bill"
        Exit Sub
    End If
    creditbill.CMDEXIT.Enabled = False
    If optCash.Value = False Then
        creditbill.lblcredit.Caption = "1"
        creditbill.GRDRECEIPT.rows = 1
    Else
        creditbill.lblcredit.Caption = "0"
        creditbill.GRDRECEIPT.rows = 1
        creditbill.GRDRECEIPT.TextMatrix(0, 0) = Val(txtrcptamt.Text)
        creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
        creditbill.GRDRECEIPT.TextMatrix(1, 0) = Trim(TXTREFNO.Text)
        If OptBank.Value = True Then
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(2, 0) = "B"
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(3, 0) = Trim(TxtChqNo.Text)
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(4, 0) = CMBDISTI.BoundText
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(5, 0) = Format(DtChqDate.Value, "DD/MM/YYYY")
            If ChkStatus.Value = 0 Then
                creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
                creditbill.GRDRECEIPT.TextMatrix(6, 0) = "N"
            Else
                creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
                creditbill.GRDRECEIPT.TextMatrix(6, 0) = "Y"
            End If
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(7, 0) = CMBDISTI.Text
        Else
            creditbill.GRDRECEIPT.rows = creditbill.GRDRECEIPT.rows + 1
            creditbill.GRDRECEIPT.TextMatrix(2, 0) = "C"
        End If
        
    End If
    creditbill.Enabled = True
    'On Error Resume Next
    If LBLTYPE.Caption = "S" Then
        Call creditbill.AppendSale
    Else
        Call creditbill.Generateprint
    End If
    Unload Me
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Form_Load()
    cetre Me
    AGNT_FLAG = True
    Call fillcombo
    If creditbill.GRDRECEIPT.rows > 1 And Val(creditbill.GRDRECEIPT.TextMatrix(0, 0)) > 0 Then
        optCash.Value = True
        txtrcptamt.Text = Val(creditbill.GRDRECEIPT.TextMatrix(0, 0))
        TXTREFNO.Text = Trim(creditbill.GRDRECEIPT.TextMatrix(1, 0))
        If creditbill.GRDRECEIPT.rows > 2 Then
            If creditbill.GRDRECEIPT.TextMatrix(2, 0) = "B" Then
                OptBank.Value = True
                TxtChqNo.Text = Trim(creditbill.GRDRECEIPT.TextMatrix(3, 0))
                CMBDISTI.BoundText = creditbill.GRDRECEIPT.TextMatrix(4, 0)
                If IsDate(creditbill.GRDRECEIPT.TextMatrix(5, 0)) Then
                    DtChqDate.Value = Format(creditbill.GRDRECEIPT.TextMatrix(5, 0), "DD/MM/YYYY")
                Else
                    DtChqDate.Value = Format(creditbill.TXTINVDATE.Text, "DD/MM/YYYY")
                End If
                If creditbill.GRDRECEIPT.TextMatrix(6, 0) = "N" Then
                    ChkStatus.Value = 0
                Else
                    ChkStatus.Value = 1
                End If
                'creditbill.GRDRECEIPT.TextMatrix(7, 0) = CMBDISTI.Text
            Else
                Opt_Cash.Value = True
            End If
        Else
            optCash.Value = True
        End If
    Else
       OptCredit.Value = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
    creditbill.Enabled = True
End Sub

Private Sub optCash_Click()
    FRMEMAIN.Visible = True
    If Val(txtrcptamt.Text) = 0 Then
        txtrcptamt.Text = Val(creditbill.lblnetamount.Caption)
    End If
End Sub

Private Sub OptCredit_Click()
    FRMEMAIN.Visible = False
End Sub

Private Sub OptBank_Click()
    FrmBank.Visible = True
End Sub

Private Sub Opt_Cash_Click()
    FrmBank.Visible = False
End Sub

Private Sub txtrcptamt_GotFocus()
    txtrcptamt.SelStart = 0
    txtrcptamt.SelLength = Len(txtrcptamt.Text)
End Sub

Private Sub txtrcptamt_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(txtrcptamt.Text) = 0 Then
                MsgBox "Enter Payment Amount", vbOKOnly, "PAYMENT..."
                txtrcptamt.SetFocus
                Exit Sub
            End If
            TXTREFNO.SetFocus
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
    TXTREFNO.SelLength = Len(TXTREFNO.Text)
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


