VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmINVTYPE 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   3240
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMINVTYPE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3240
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   3075
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4230
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   555
         Left            =   705
         TabIndex        =   0
         Top             =   2370
         Width           =   1410
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Cancel"
         Height          =   555
         Left            =   2175
         TabIndex        =   1
         Top             =   2370
         Width           =   1380
      End
      Begin MSForms.OptionButton OptPetty 
         Height          =   495
         Left            =   525
         TabIndex        =   5
         Top             =   1680
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
         Top             =   435
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
         Top             =   1065
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
Private Sub CMDEXIT_Click()
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    creditbill.cmdexit.Enabled = False
    creditbill.Enabled = True
    If Opt8B.value = True Then
        Call creditbill.Make_Invoice("GI")
    ElseIf Opt8.value = True Then
        Call creditbill.Make_Invoice("HI")
    Else
        Call creditbill.Make_Invoice("WO")
    End If
    
    Unload Me
End Sub

Private Sub Form_Load()
    cetre Me
    Opt8B.value = True
End Sub
