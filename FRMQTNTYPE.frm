VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmQTNTYPE 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMQTNTYPE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3000
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   2805
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4230
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   555
         Left            =   705
         TabIndex        =   0
         Top             =   2145
         Width           =   1410
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Cancel"
         Height          =   555
         Left            =   2175
         TabIndex        =   1
         Top             =   2145
         Width           =   1380
      End
      Begin MSForms.OptionButton OptB2C 
         Height          =   495
         Left            =   1320
         TabIndex        =   6
         Top             =   840
         Width           =   2295
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "4048;873"
         Value           =   "1"
         Caption         =   "B2C BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton Opt8B 
         Height          =   495
         Left            =   1320
         TabIndex        =   5
         Top             =   240
         Width           =   2295
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "4048;873"
         Value           =   "0"
         Caption         =   "GST BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton Opt8 
         Height          =   495
         Left            =   1305
         TabIndex        =   4
         Top             =   2715
         Visible         =   0   'False
         Width           =   1965
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "3466;873"
         Value           =   "0"
         Caption         =   "8 BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optpetty 
         Height          =   495
         Left            =   1320
         TabIndex        =   3
         Top             =   1485
         Width           =   1965
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   64
         DisplayStyle    =   5
         Size            =   "3466;873"
         Value           =   "0"
         Caption         =   "PETTY BILL"
         FontName        =   "Arial Rounded MT Bold"
         FontHeight      =   405
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
End
Attribute VB_Name = "frmQTNTYPE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdexit_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If OptPetty.Value = True Then
        creditbill.lblcredit.Caption = "2"
    ElseIf OptB2C.Value = True Then
        creditbill.lblcredit.Caption = "1"
    Else
        creditbill.lblcredit.Caption = "0"
    End If
    creditbill.CMDEXIT.Enabled = False
    creditbill.Enabled = True
    Call creditbill.Make_Invoice
    
    Unload Me
End Sub

Private Sub Form_Load()
    cetre Me
    If creditbill.lblcredit.Caption = "2" Then
        OptPetty.Value = True
    ElseIf creditbill.lblcredit.Caption = "1" Then
        OptB2C.Value = True
    Else
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
            Opt8B.Visible = False
            OptB2C.Value = True
        Else
            Opt8B.Value = True
        End If
    End If
End Sub

