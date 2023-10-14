VERSION 5.00
Begin VB.Form frmabout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4680
   Icon            =   "frmabout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmabout.frx":0442
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   840
      Left            =   15
      TabIndex        =   4
      Top             =   1530
      Width           =   4665
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Service Support"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   45
         TabIndex        =   6
         Top             =   120
         Width           =   3990
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "9745231707, 7902608968, 9072999926"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   60
         TabIndex        =   5
         Top             =   465
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   15
      TabIndex        =   1
      Top             =   2325
      Width           =   4665
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Sales Enquiry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   45
         TabIndex        =   3
         Top             =   150
         Width           =   3990
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "8139006575, 9495618968, 9072999926"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   360
         Left            =   60
         TabIndex        =   2
         Top             =   450
         Width           =   4395
      End
   End
   Begin VB.Line Line1 
      X1              =   60
      X2              =   4575
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Version 15.1.1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   15
      TabIndex        =   0
      Top             =   1110
      Width           =   4095
   End
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = 0
End Sub

Private Sub Form_LostFocus()
    Unload Me
End Sub
