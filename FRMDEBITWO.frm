VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRMDEBITWO 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMDEBITWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   1695
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   4230
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   1485
         TabIndex        =   0
         Top             =   1050
         Width           =   1185
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "PAYMENT MODE"
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
         Height          =   1410
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   4005
         Begin VB.CommandButton CMDEXIT 
            Caption         =   "&CANCEL"
            Height          =   405
            Left            =   2790
            TabIndex        =   5
            Top             =   870
            Width           =   1155
         End
         Begin MSForms.OptionButton OPTCASH 
            Height          =   300
            Left            =   480
            TabIndex        =   4
            Top             =   300
            Width           =   1365
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2408;529"
            Value           =   "0"
            Caption         =   "CASH"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton OPTCREDIT 
            Height          =   300
            Left            =   2100
            TabIndex        =   3
            Top             =   300
            Width           =   1140
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2011;529"
            Value           =   "1"
            Caption         =   "CREDIT"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
   End
End
Attribute VB_Name = "FRMDEBITWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
    FRMWITHOUT.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If OPTCASH.Value = True Then
        FRMWITHOUT.lblcredit.Caption = "0"
    Else
        FRMWITHOUT.lblcredit.Caption = "1"
    End If
    FRMWITHOUT.Enabled = True
    FRMWITHOUT.Generateprint
    Unload Me
End Sub

Private Sub Form_Load()
    cetre Me
    If FRMWITHOUT.lblcredit.Caption = "0" Then
        OPTCASH.Value = True
    Else
       OPTCREDIT.Value = True
    End If
End Sub
