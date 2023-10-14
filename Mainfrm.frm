VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Charge Sheeted Cases"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   DrawStyle       =   2  'Dot
   FillStyle       =   0  'Solid
   Icon            =   "Mainfrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmsCancel 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      TabIndex        =   27
      Top             =   8280
      Width           =   1815
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9720
      MaskColor       =   &H8000000A&
      TabIndex        =   26
      Top             =   7320
      Width           =   1815
   End
   Begin VB.TextBox TxtRemarks 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   40
      TabIndex        =   24
      Top             =   8640
      Width           =   5775
   End
   Begin VB.TextBox TxtAddlCase 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   23
      Top             =   7680
      Width           =   5775
   End
   Begin VB.TextBox TxtCasedetails 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   6720
      Width           =   5775
   End
   Begin VB.TextBox TxtDistrict 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   20
      TabIndex        =   21
      Top             =   6120
      Width           =   3255
   End
   Begin VB.TextBox TxtAddress 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3360
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4800
      Width           =   5775
   End
   Begin VB.TextBox TxtNationality 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   15
      TabIndex        =   19
      Top             =   4200
      Width           =   3255
   End
   Begin VB.TextBox TxtNamewithfath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   50
      TabIndex        =   17
      Top             =   2760
      Width           =   5655
   End
   Begin VB.TextBox TxtFather 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   16
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox TxtAlias 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   15
      Top             =   1560
      Width           =   3255
   End
   Begin VB.TextBox TxtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3360
      MaxLength       =   25
      TabIndex        =   14
      Top             =   960
      Width           =   3255
   End
   Begin MSMask.MaskEdBox MaskEdBirthdate 
      Height          =   420
      Left            =   3360
      TabIndex        =   18
      Top             =   3600
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox MaskEdDate 
      Height          =   420
      Left            =   10080
      TabIndex        =   25
      Top             =   960
      Width           =   1440
      _ExtentX        =   2540
      _ExtentY        =   741
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "dd/mm/yyyy"
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.Label Lblrecord 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   10080
      TabIndex        =   28
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Label Label15 
      Caption         =   "Date of entry"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7440
      TabIndex        =   13
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label14 
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   12
      Top             =   8640
      Width           =   3015
   End
   Begin VB.Label Label13 
      Caption         =   "Addl Case Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   11
      Top             =   7920
      Width           =   2895
   End
   Begin VB.Label Label12 
      Caption         =   "Record No"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   1680
      Width           =   1815
   End
   Begin VB.Label Label11 
      Caption         =   "Case Details"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   7080
      Width           =   2535
   End
   Begin VB.Label Label10 
      Caption         =   "District"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   8
      Top             =   6240
      Width           =   3015
   End
   Begin VB.Label Label9 
      Caption         =   "Nationality"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   4200
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   5160
      Width           =   3015
   End
   Begin VB.Label Label7 
      Caption         =   "Date of Birth"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Name of Father"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Name Appended With Father's name"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   3015
   End
   Begin VB.Label Label3 
      Caption         =   "Alias Name"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Details of Persons Involved in Charge Sheeted Crime Cases"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdSave_Click()
    If TxtName.Text = "" Then
        MsgBox "Name cannot be blank", vbInformation, "Save"
        TxtName.SetFocus
    Else
        TxtName.Text = ""
        TxtAlias.Text = ""
        TxtNamewithfath.Text = ""
        TxtFather.Text = ""
        MaskEdBirthdate.Text = "  /  /    "
        TxtNationality.Text = ""
        TxtAddress.Text = ""
        TxtDistrict.Text = ""
        TxtCasedetails.Text = ""
        TxtAddlCase.Text = ""
        TxtRemarks.Text = ""
    End If
    
End Sub

Private Sub CmsCancel_Click()
Unload Me
End Sub

Private Sub Form_Load()
MaskEdDate.Text = Date
TxtDistrict = "Alappuzha"
TxtNationality = "Indian"
End Sub

Private Sub MaskEdBirthdate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
             TxtNationality.SelStart = 0
             TxtNationality.SelLength = Len(TxtNationality.Text)
             TxtNationality.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub MaskEdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           CmdSave.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtAddlCase_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtRemarks.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
             TxtDistrict.SelStart = 0
             TxtDistrict.SelLength = Len(TxtDistrict.Text)
             TxtDistrict.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtAlias_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
            TxtFather.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtBirthdate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtAlias.SetFocus
        Case vbKeyUp
            
    End Select
End Sub


Private Sub TxtCasedetails_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtAddlCase.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtDistrict_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtCasedetails.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtFather_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtNamewithfath.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtAlias.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtNamewithfath_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           MaskEdBirthdate.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtNationality_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
           TxtAddress.SetFocus
        Case vbKeyUp
            
    End Select
End Sub

Private Sub TxtRemarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
            MaskEdDate.SelStart = 0
            MaskEdDate.SelLength = Len(MaskEdDate.Text)
            MaskEdDate.SetFocus
        Case vbKeyUp
            
    End Select
End Sub
