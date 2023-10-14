VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmCREDIT 
   BorderStyle     =   0  'None
   Caption         =   "Set Minimum Stock"
   ClientHeight    =   1980
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   4350
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMCREDITWO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame 
      Height          =   1695
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   4230
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
         Height          =   750
         Left            =   390
         TabIndex        =   3
         Top             =   225
         Width           =   3480
         Begin MSForms.OptionButton OPTCASH 
            Height          =   300
            Left            =   120
            TabIndex        =   5
            Top             =   270
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
            Left            =   2025
            TabIndex        =   4
            Top             =   270
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
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   810
         TabIndex        =   0
         Top             =   1050
         Width           =   1200
      End
      Begin VB.CommandButton cmdexit 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2115
         TabIndex        =   1
         Top             =   1050
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCREDIT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMDEXIT_Click()
    MDIMAIN.cmdpurchase.Enabled = True
    creditbill.Enabled = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If OPTCASH.value = True Then
        creditbill.lblcredit.Caption = "0"
    Else
        creditbill.lblcredit.Caption = "1"
    End If
    MDIMAIN.cmdpurchase.Enabled = True
    creditbill.Enabled = True
    creditbill.appendpurchase
    Unload Me
End Sub

Private Sub Form_Load()
    cetre Me
    If creditbill.lblcredit.Caption = "0" Then
        OPTCASH.value = True
    Else
       OPTCREDIT.value = True
    End If
End Sub

