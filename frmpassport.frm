VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmpassport1 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "passport"
   ClientHeight    =   7290
   ClientLeft      =   1095
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmpassport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmpassport.frx":030A
   Moveable        =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   6495
      Left            =   10680
      TabIndex        =   36
      Top             =   120
      Width           =   1695
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   540
         Left            =   240
         TabIndex        =   15
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   540
         Left            =   240
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   540
         Left            =   240
         TabIndex        =   16
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   540
         Left            =   240
         TabIndex        =   13
         Top             =   2280
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   540
         Left            =   240
         TabIndex        =   12
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
         Height          =   540
         Left            =   240
         TabIndex        =   37
         Top             =   1560
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   540
         Left            =   240
         TabIndex        =   38
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Station Details"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   31
      Top             =   120
      Width           =   10575
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1920
         MaxLength       =   20
         TabIndex        =   34
         Top             =   480
         Width           =   4695
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   7920
         MaxLength       =   6
         TabIndex        =   33
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "GL NO"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   2
         Left            =   6960
         TabIndex        =   35
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Police Station"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Frame FrmCase 
      BackColor       =   &H00C0E0FF&
      Caption         =   "  Charge Sheeted Cases  "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5295
      Left            =   120
      TabIndex        =   17
      Top             =   1200
      Width           =   10575
      Begin VB.TextBox TxtAge 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   2
         TabIndex        =   3
         Top             =   1320
         Width           =   735
      End
      Begin VB.PictureBox picStatBox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2040
         ScaleHeight     =   300
         ScaleWidth      =   4875
         TabIndex        =   29
         Top             =   5880
         Width           =   4875
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "FNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   0
         Top             =   600
         Width           =   8175
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "LNAME"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   25
         TabIndex        =   1
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "PRNTGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   6840
         MaxLength       =   25
         TabIndex        =   2
         Top             =   960
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "NATIONAL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   6840
         MaxLength       =   10
         TabIndex        =   5
         Top             =   1320
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "PRSNT_ADD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   5
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "DIST1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   6840
         MaxLength       =   15
         TabIndex        =   7
         Top             =   1680
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "INFRMTN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   7
         Left            =   2040
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "REFCE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   8
         Left            =   6840
         MaxLength       =   50
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "REMRKS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Index           =   9
         Left            =   2040
         MaxLength       =   25
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3360
         Width           =   3375
      End
      Begin VB.TextBox txtFields 
         Appearance      =   0  'Flat
         DataField       =   "REK"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   6840
         MaxLength       =   5
         TabIndex        =   11
         Top             =   3480
         Width           =   975
      End
      Begin MSMask.MaskEdBox MaskEdDate 
         Height          =   300
         Left            =   3600
         TabIndex        =   4
         Top             =   1320
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   8
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
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Record No"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   5520
         TabIndex        =   30
         Top             =   3480
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   28
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   27
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Father"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5520
         TabIndex        =   25
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Age"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "District"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5520
         TabIndex        =   22
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Alias Name"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Nationality"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Case Details"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Addl Case Details"
         BeginProperty Font 
            Name            =   "Arial Rounded MT Bold"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   5520
         TabIndex        =   18
         Top             =   2520
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   39
      Top             =   6480
      Width           =   12375
      Begin VB.CommandButton cmdLast 
         Height          =   300
         Left            =   7665
         Picture         =   "frmpassport.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdNext 
         Height          =   300
         Left            =   7320
         Picture         =   "frmpassport.frx":0956
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdPrevious 
         Height          =   300
         Left            =   3465
         Picture         =   "frmpassport.frx":0C98
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.CommandButton cmdFirst 
         Height          =   300
         Left            =   3120
         Picture         =   "frmpassport.frx":0FDA
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   345
      End
      Begin VB.Label lblStatus 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   3810
         TabIndex        =   44
         Top             =   360
         Width           =   3360
      End
   End
End
Attribute VB_Name = "frmpassport1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean

Private Sub Form_Load()
  Dim db As Connection
  Set db = New Connection
  db.CursorLocation = adUseClient
  'db.Open "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=D:\Passport Verification\passport1.mdb;"
  db.Open "PROVIDER=Microsoft.Jet.OLEDB.4.0;Data Source=" + App.Path + "\passport.mdb;"

  Set adoPrimaryRS = New Recordset
  adoPrimaryRS.Open "select FNAME,LNAME,PRNTGE,D_B,NATIONAL,PRSNT_ADD,DIST1,INFRMTN,REFCE,REMRKS,DT_OF_INDX,REK from passport", db, adOpenStatic, adLockOptimistic

  Dim oText As TextBox
  'Bind the text boxes to the data provider
  For Each oText In Me.txtFields
    Set oText.DataSource = adoPrimaryRS
  Next
    If Not (adoPrimaryRS.BOF And adoPrimaryRS.EOF) Then
        If adoPrimaryRS.Fields!D_B <> Null Then
            MaskEdDate.Text = adoPrimaryRS.Fields!D_B
            TxtAge.Text = Val(Year(Date)) - Val(Mid(MaskEdDate.Text, 7, 4))
        End If
    End If
  mbDataChanged = False
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If mbEditFlag Or mbAddNewFlag Then Exit Sub

  Select Case KeyCode
    Case vbKeyEscape
      cmdClose_Click
    Case vbKeyEnd
      cmdLast_Click
    Case vbKeyHome
      cmdFirst_Click
    Case vbKeyUp, vbKeyPageUp
      If Shift = vbCtrlMask Then
        cmdFirst_Click
      Else
        cmdPrevious_Click
      End If
    Case vbKeyDown, vbKeyPageDown
      If Shift = vbCtrlMask Then
        cmdLast_Click
      Else
        cmdNext_Click
      End If
  End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Screen.MousePointer = vbDefault
End Sub

'Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This will display the current record position for this recordset
 ' lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
'End Sub

Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
  'This is where you put validation code
  'This event gets called when the following actions occur
  Dim bCancel As Boolean

  Select Case adReason
  Case adRsnAddNew
  Case adRsnClose
  Case adRsnDelete
  Case adRsnFirstChange
  Case adRsnMove
  Case adRsnRequery
  Case adRsnResynch
  Case adRsnUndoAddNew
  Case adRsnUndoDelete
  Case adRsnUndoUpdate
  Case adRsnUpdate
  End Select

  If bCancel Then adStatus = adStatusCancel
End Sub

Private Sub cmdAdd_Click()
  On Error GoTo AddErr
    
    
    SetButtons True
    mbEditFlag = False
    mbAddNewFlag = False
 
  With adoPrimaryRS
    If Not (.BOF And .EOF) Then
      mvBookMark = .Bookmark
    End If
    .AddNew
    
    lblStatus.Caption = "Add record"
    mbAddNewFlag = True
    SetButtons False
  End With
    FrmCase.Enabled = True
    txtFields(4) = "Indian"
    txtFields(6) = "Alappuzha"
    MaskEdDate.Text = "  /  /    "
    TxtAge.Text = ""
    txtFields(0).SetFocus
    
  Exit Sub
AddErr:
  MsgBox err.Description
End Sub

Private Sub cmdDelete_Click()
  On Error GoTo DeleteErr
  With adoPrimaryRS
    .Delete
    .MoveNext
    If .EOF Then .MoveLast
  End With
  Exit Sub
DeleteErr:
  MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
  'This is only needed for multi user apps
  On Error GoTo RefreshErr
  adoPrimaryRS.Requery
  Exit Sub
RefreshErr:
  MsgBox err.Description
End Sub

Private Sub cmdEdit_Click()
  On Error GoTo EditErr

    FrmCase.Enabled = True
    txtFields(0).SetFocus

  lblStatus.Caption = "Edit record"
  mbEditFlag = True
  SetButtons False
  Exit Sub

EditErr:
  MsgBox err.Description
End Sub
Private Sub cmdCancel_Click()
  On Error Resume Next

  SetButtons True
  mbEditFlag = False
  mbAddNewFlag = False
  adoPrimaryRS.CancelUpdate
  If mvBookMark > 0 Then
    adoPrimaryRS.Bookmark = mvBookMark
  Else
    adoPrimaryRS.MoveFirst
  End If
  mbDataChanged = False
  FrmCase.Enabled = False
  txtFields(4) = ""
  txtFields(6) = ""

End Sub

Private Sub cmdUpdate_Click()
    
  On Error GoTo UpdateErr
    
  adoPrimaryRS.UpdateBatch adAffectAll
  

  If mbAddNewFlag Then
    adoPrimaryRS.MoveLast              'move to the new record
    
    If Not (MaskEdDate.Text = "  /  /    " Or Trim(Mid(MaskEdDate.Text, 1, 6)) = "00/00/") Then
       If Not IsDate(MaskEdDate.Text) Then
            MsgBox "Enter Date of Birth Properly."
            MaskEdDate.SetFocus
            Exit Sub
        End If
    End If
        
    adoPrimaryRS.Fields!D_B = MaskEdDate.Text
    adoPrimaryRS.Fields!DT_OF_INDX = Date
  End If

  mbEditFlag = False
  mbAddNewFlag = False
  SetButtons True
  mbDataChanged = False

  Exit Sub

UpdateErr:
  MsgBox err.Description
End Sub

Private Sub cmdClose_Click()
  Unload Me
End Sub

Private Sub cmdFirst_Click()
  On Error GoTo GoFirstError

  adoPrimaryRS.MoveFirst
  mbDataChanged = False

  Exit Sub

GoFirstError:
  MsgBox err.Description
End Sub

Private Sub cmdLast_Click()
  On Error GoTo GoLastError

  adoPrimaryRS.MoveLast
  mbDataChanged = False

  Exit Sub

GoLastError:
  MsgBox err.Description
End Sub

Private Sub cmdNext_Click()
  On Error GoTo GoNextError
    
    MaskEdDate.Text = "  /  /  "
    TxtAge.Text = ""
    If Not adoPrimaryRS.EOF Or adoPrimaryRS.BOF Then
        adoPrimaryRS.MoveNext
                MaskEdDate.Text = adoPrimaryRS.Fields!D_B
                TxtAge.Text = Val(Year(Date)) - Val(Mid(MaskEdDate.Text, 7, 4))
       
    End If
  If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
     'moved off the end so go back
    adoPrimaryRS.MoveLast
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub
GoNextError:
  MsgBox err.Description
End Sub

Private Sub cmdPrevious_Click()
  On Error GoTo GoPrevError
    
    MaskEdDate.Text = "  /  /    "
    TxtAge.Text = ""
  If Not adoPrimaryRS.BOF Then
    adoPrimaryRS.MovePrevious
    
    If Not adoPrimaryRS.BOF Then
        If adoPrimaryRS.Fields!D_B <> Null Then
            MaskEdDate.Text = adoPrimaryRS.Fields!D_B
            TxtAge.Text = Val(Year(Date)) - Val(Mid(MaskEdDate.Text, 7, 4))
        End If
    End If
        
    End If
  If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
    Beep
    'moved off the end so go back
    adoPrimaryRS.MoveFirst
  End If
  'show the current record
  mbDataChanged = False

  Exit Sub

GoPrevError:
  MsgBox err.Description
End Sub

Private Sub SetButtons(bVal As Boolean)
  cmdAdd.Visible = bVal
  cmdEdit.Visible = bVal
  cmdUpdate.Visible = Not bVal
  cmdCancel.Visible = Not bVal
  cmddelete.Visible = bVal
  cmdclose.Visible = bVal
  cmdRefresh.Visible = bVal
  cmdnext.Enabled = bVal
  cmdFirst.Enabled = bVal
  cmdLast.Enabled = bVal
  cmdPrevious.Enabled = bVal
End Sub


Private Sub MaskEdDate_GotFocus()
        MaskEdDate.SelStart = 0
        MaskEdDate.SelLength = Len(MaskEdDate.Text)
End Sub

Private Sub MaskEdDate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
                txtFields(4).SetFocus
        Case vbKeyUp
    End Select
End Sub

Private Sub picButtons_Click()

End Sub

Private Sub TxtAge_GotFocus()
    TxtAge.SelStart = 0
    TxtAge.SelLength = Len(TxtAge.Text)
End Sub

Private Sub TxtAge_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
                MaskEdDate.SetFocus
        Case vbKeyUp
    End Select
End Sub

Private Sub TxtAge_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtAge_LostFocus()
    If Val(TxtAge.Text) = 0 Then Exit Sub
    'MaskEdDate.Text = Trim(("00/00/") & (Val(Year(Date)) - Val(TxtAge.Text)))
    
End Sub

Private Sub txtFields_GotFocus(Index As Integer)
            
        txtFields(Index).SelStart = 0
        txtFields(Index).SelLength = Len(txtFields(Index).Text)
        
End Sub

Private Sub txtFields_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyTab, vbKeyReturn, vbKeyDown
            If Index >= 9 Then
                cmdUpdate.SetFocus
                Exit Sub
            End If
     
           If Index = 2 Then
                Me.TxtAge.SetFocus
                   Else
                   txtFields(Index + 1).SetFocus
            End If
        Case vbKeyUp
            
    End Select
End Sub

Private Sub txtFields_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index <> 3 Then Exit Sub
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeySeparator, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtFields_LostFocus(Index As Integer)
    If Index = 2 Then txtFields(0).Text = txtFields(0) & " " & txtFields(2)
End Sub
