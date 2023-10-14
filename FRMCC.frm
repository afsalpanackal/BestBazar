VERSION 5.00
Begin VB.Form FrmCC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cost Code"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   Icon            =   "FRMCC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4095
      TabIndex        =   10
      Top             =   1770
      Width           =   1215
   End
   Begin VB.TextBox txt9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2400
      Width           =   1005
   End
   Begin VB.TextBox Txt8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   8
      Top             =   1815
      Width           =   1005
   End
   Begin VB.TextBox Txt7 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1245
      Width           =   1005
   End
   Begin VB.TextBox Txt6 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   6
      Top             =   675
      Width           =   1005
   End
   Begin VB.TextBox Txt5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   2790
      MaxLength       =   1
      TabIndex        =   5
      Top             =   120
      Width           =   1005
   End
   Begin VB.TextBox Txt4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   4
      Top             =   2400
      Width           =   1005
   End
   Begin VB.TextBox Txt3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1815
      Width           =   1005
   End
   Begin VB.TextBox Txt2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   2
      Top             =   1245
      Width           =   1005
   End
   Begin VB.TextBox Txt1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   1
      Top             =   675
      Width           =   1005
   End
   Begin VB.TextBox Txt0 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1020
      MaxLength       =   1
      TabIndex        =   0
      Top             =   120
      Width           =   1005
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&XIT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   4095
      TabIndex        =   11
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   9
      Left            =   1980
      TabIndex        =   21
      Top             =   2460
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   8
      Left            =   1980
      TabIndex        =   20
      Top             =   1920
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   7
      Left            =   2010
      TabIndex        =   19
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   6
      Left            =   2010
      TabIndex        =   18
      Top             =   750
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   5
      Left            =   2010
      TabIndex        =   17
      Top             =   150
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   4
      Left            =   240
      TabIndex        =   16
      Top             =   2460
      Width           =   500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   3
      Left            =   255
      TabIndex        =   15
      Top             =   1920
      Width           =   500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   2
      Left            =   255
      TabIndex        =   14
      Top             =   1320
      Width           =   500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   1
      Left            =   255
      TabIndex        =   13
      Top             =   750
      Width           =   500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   0
      Left            =   255
      TabIndex        =   12
      Top             =   150
      Width           =   500
   End
End
Attribute VB_Name = "FrmCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    
    If Trim(Txt0.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt0.SetFocus
        Exit Sub
    End If
    If Trim(Txt1.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt1.SetFocus
        Exit Sub
    End If
    If Trim(Txt2.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt2.SetFocus
        Exit Sub
    End If
    If Trim(Txt3.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt3.SetFocus
        Exit Sub
    End If
    If Trim(Txt4.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt4.SetFocus
        Exit Sub
    End If
    If Trim(Txt5.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt5.SetFocus
        Exit Sub
    End If
    If Trim(Txt6.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt6.SetFocus
        Exit Sub
    End If
    If Trim(Txt7.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt7.SetFocus
        Exit Sub
    End If
    If Trim(Txt8.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        Txt8.SetFocus
        Exit Sub
    End If
    If Trim(txt9.Text) = "" Then
        MsgBox "Field cannot be empty", , "EzBiz"
        txt9.SetFocus
        Exit Sub
    End If
    
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo eRRhAND
    db.Execute "delete from ccode"
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM ccode ", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        RSTCOMPANY.AddNew
    End If
    RSTCOMPANY!CC0 = Trim(Txt0.Text)
    RSTCOMPANY!CC1 = Trim(Txt1.Text)
    RSTCOMPANY!CC2 = Trim(Txt2.Text)
    RSTCOMPANY!CC3 = Trim(Txt3.Text)
    RSTCOMPANY!CC4 = Trim(Txt4.Text)
    RSTCOMPANY!CC5 = Trim(Txt5.Text)
    RSTCOMPANY!CC6 = Trim(Txt6.Text)
    RSTCOMPANY!CC7 = Trim(Txt7.Text)
    RSTCOMPANY!CC8 = Trim(Txt8.Text)
    RSTCOMPANY!CC9 = Trim(txt9.Text)
    RSTCOMPANY.Update

    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    MsgBox "Saved successfully", , "EzBiz"
    Exit Sub
eRRhAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Form_Load()
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo eRRhAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM ccode ", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Txt0.Text = IIf(IsNull(RSTCOMPANY!CC0), "", RSTCOMPANY!CC0)
        Txt1.Text = IIf(IsNull(RSTCOMPANY!CC1), "", RSTCOMPANY!CC1)
        Txt2.Text = IIf(IsNull(RSTCOMPANY!CC2), "", RSTCOMPANY!CC2)
        Txt3.Text = IIf(IsNull(RSTCOMPANY!CC3), "", RSTCOMPANY!CC3)
        Txt4.Text = IIf(IsNull(RSTCOMPANY!CC4), "", RSTCOMPANY!CC4)
        Txt5.Text = IIf(IsNull(RSTCOMPANY!CC5), "", RSTCOMPANY!CC5)
        Txt6.Text = IIf(IsNull(RSTCOMPANY!CC6), "", RSTCOMPANY!CC6)
        Txt7.Text = IIf(IsNull(RSTCOMPANY!CC7), "", RSTCOMPANY!CC7)
        Txt8.Text = IIf(IsNull(RSTCOMPANY!CC8), "", RSTCOMPANY!CC8)
        txt9.Text = IIf(IsNull(RSTCOMPANY!CC9), "", RSTCOMPANY!CC9)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Exit Sub
eRRhAND:
    MsgBox err.Description, , "EzBiz"
End Sub
