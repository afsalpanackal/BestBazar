VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmDUPbilltype 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDUPbilltype.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   1695
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   4230
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "MODEL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   390
         TabIndex        =   5
         Top             =   195
         Width           =   3480
         Begin MSForms.OptionButton OPTSLIP 
            Height          =   300
            Left            =   120
            TabIndex        =   1
            Top             =   330
            Width           =   1365
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2408;529"
            Value           =   "1"
            Caption         =   "SLIP"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton OPTBILL 
            Height          =   300
            Left            =   1995
            TabIndex        =   2
            Top             =   330
            Width           =   1275
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2249;529"
            Value           =   "0"
            Caption         =   "BILL"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   810
         TabIndex        =   0
         Top             =   1140
         Width           =   1200
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2115
         TabIndex        =   3
         Top             =   1140
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmDUPbilltype"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    FRMDUPLI.Enabled = True
    Unload Me
End Sub

Private Sub cmdexit_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyS, Asc("s")
            OPTSLIP.SetFocus
        Case vbKeyB, Asc("b")
            OPTBILL.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
End Sub

Private Sub cmdOK_Click()
    If OPTSLIP.Value = True Then
        FRMDUPLI.PRINTSLIP
    Else
        If FRMDUPLI.OPTAUTOMATIC = True Then
            FRMDUPLI.AutoPRINTBILL
        Else
            FRMDUPLI.PRINTBILL
        End If
    End If
    FRMDUPLI.Enabled = True
    Unload Me
    FRMDUPLI.cmdRefresh.SetFocus
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyS, Asc("s")
            OPTSLIP.SetFocus
        Case vbKeyB, Asc("b")
            OPTBILL.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
End Sub

Private Sub Form_Load()
    cetre Me
End Sub

Private Sub OPTBILL_GotFocus()
    OPTBILL.Value = True
End Sub

Private Sub OPTBILL_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call cmdOK_Click
        Case vbKeyEscape
            Call cmdexit_Click
        Case vbKeyS, Asc("s")
            OPTSLIP.SetFocus
        Case vbKeyB, Asc("b")
            OPTBILL.SetFocus
    End Select
End Sub

Private Sub OPTSLIP_GotFocus()
    OPTSLIP.Value = True
End Sub

Private Sub OPTSLIP_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call cmdOK_Click
        Case vbKeyEscape
            Call cmdexit_Click
        Case vbKeyS, Asc("s")
            OPTSLIP.SetFocus
        Case vbKeyB, Asc("b")
            OPTBILL.SetFocus
    End Select
End Sub
