VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmINVTYPE 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3465
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMINVTYPE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3465
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   3435
      Left            =   45
      TabIndex        =   2
      Top             =   -15
      Width           =   4230
      Begin VB.Frame FrmYear 
         Caption         =   "Financial Year"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   1275
         TabIndex        =   6
         Top             =   165
         Visible         =   0   'False
         Width           =   1935
         Begin VB.TextBox txtyear 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   135
            MaxLength       =   4
            TabIndex        =   7
            Top             =   270
            Width           =   1650
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   555
         Left            =   705
         TabIndex        =   0
         Top             =   2790
         Width           =   1410
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Cancel"
         Height          =   555
         Left            =   2175
         TabIndex        =   1
         Top             =   2775
         Width           =   1380
      End
      Begin MSForms.OptionButton OptPetty 
         Height          =   495
         Left            =   525
         TabIndex        =   5
         Top             =   2175
         Width           =   3660
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "6456;873"
         Value           =   "0"
         Caption         =   "PETTY SALES"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton Opt8B 
         Height          =   495
         Left            =   525
         TabIndex        =   4
         Top             =   1065
         Width           =   3720
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "6562;873"
         Value           =   "1"
         Caption         =   "GST (B2B) BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton Opt8 
         Height          =   495
         Left            =   525
         TabIndex        =   3
         Top             =   1635
         Width           =   3660
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "6456;873"
         Value           =   "0"
         Caption         =   "GST (B2C) BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmINVTYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdExit_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If FrmYear.Visible = True Then
        If Len(txtyear.text) <> 4 Then
            MsgBox "Please enter proper year", vbOKOnly, "EzBiz"
            txtyear.SetFocus
            Exit Sub
        End If
        
        If Val(txtyear.text) < 2010 Or Val(txtyear.text) > 2030 Then
            MsgBox "Unexpected error occured", vbOKOnly, "EzBiz"
            txtyear.SetFocus
            Exit Sub
        End If
    End If

    creditbill.CMDEXIT.Enabled = False
    creditbill.Enabled = True
    If Opt8B.Value = True Then
        Call creditbill.Make_Invoice("GI")
    ElseIf Opt8.Value = True Then
        Call creditbill.Make_Invoice("HI")
    Else
        Call creditbill.Make_Invoice("WO")
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    cetre Me
    Opt8B.Value = True
End Sub

Private Sub txtyear_GotFocus()
    txtyear.SelStart = 0
    txtyear.SelLength = Len(txtyear.text)
    txtyear.BackColor = &H98F3C1
End Sub

Private Sub txtyear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab, vbKeyEscape
            If txtyear.text = "" Then
                txtyear.text = Year(MDIMAIN.DTFROM.Value)
                cmdOK.SetFocus
                Exit Sub
            End If
            If Len(txtyear.text) <> 4 Then
                MsgBox "Please enter proper year", vbOKOnly, "EzBiz"
                txtyear.SetFocus
                Exit Sub
            End If
            If Val(txtyear.text) < 2010 Or Val(txtyear.text) > 2030 Then
                MsgBox "Unexpected error occured", vbOKOnly, "EzBiz"
                txtyear.SetFocus
                Exit Sub
            End If
            cmdOK.SetFocus
    End Select
End Sub

Private Sub txtyear_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtyear_LostFocus()
    txtyear.BackColor = vbWhite
End Sub

