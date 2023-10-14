VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form FRMCMPINDEX 
   Caption         =   "COMPANY INDEX"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   6915
   Begin VB.CommandButton CMDREMOVE 
      BackColor       =   &H00400000&
      Caption         =   "RE&MOVE DISTRIBUTOR"
      Height          =   465
      Left            =   2670
      MaskColor       =   &H80000007&
      TabIndex        =   5
      Top             =   5940
      UseMaskColor    =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton CMDADDDIST 
      BackColor       =   &H00400000&
      Caption         =   "ADD DIS&TRIBUTOR"
      Height          =   480
      Left            =   1335
      MaskColor       =   &H80000007&
      TabIndex        =   4
      Top             =   5940
      UseMaskColor    =   -1  'True
      Width           =   1230
   End
   Begin VB.ListBox lstmanufact 
      Appearance      =   0  'Flat
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
      Height          =   5655
      Left            =   3495
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   105
      Width           =   3330
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "&EXIT"
      Height          =   450
      Left            =   90
      TabIndex        =   0
      Top             =   5970
      Width           =   1200
   End
   Begin MSDataListLib.DataList LSTDISTI 
      Height          =   1035
      Left            =   105
      TabIndex        =   3
      Top             =   615
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   1826
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSForms.ComboBox CMBCOMPANY 
      Height          =   360
      Left            =   105
      TabIndex        =   2
      Top             =   135
      Width           =   3375
      VariousPropertyBits=   746604571
      ForeColor       =   255
      DisplayStyle    =   3
      Size            =   "5953;635"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      DropButtonStyle =   0
      BorderColor     =   255
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "FRMCMPINDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim REPFLAG As Boolean
Dim focusflag As Boolean
Dim CLOSEALL As Integer
Dim RSTDISTI  As New ADODB.Recordset
Dim RSTSTOCKIST  As New ADODB.Recordset


Private Sub CMBCOMPANY_Click()
    
    On Error GoTo ERRHAND
    Set RSTDISTI = New ADODB.Recordset
    RSTDISTI.Open "SELECT ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST RIGHT JOIN COMPINDEX ON ACTMAST.ACT_CODE = COMPINDEX.ACT_CODE WHERE ITEM_NAME = '" & CMBCOMPANY.Text & "' ORDER BY ACTMAST.ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
    Set Me.LSTDISTI.RowSource = RSTDISTI
    LSTDISTI.ListField = "ACT_NAME"
    LSTDISTI.BoundColumn = "ACT_CODE"
    
    Exit Sub
    
ERRHAND:
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim RSTCOMPANY As ADODB.Recordset
    Dim RSTSTOCKIST As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ERRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select [COMP_NAME] From COMPINDEX ORDER BY [COMP_NAME]", db2, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        CMBCOMPANY.AddItem (RSTCOMPANY!COMP_NAME)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    i = 0
    Set RSTSTOCKIST = New ADODB.Recordset
    RSTSTOCKIST.Open "SELECT ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST", db2, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTSTOCKIST.EOF
        lstmanufact.AddItem (i)
        lstmanufact.List(i) = RSTSTOCKIST!ACT_NAME
        RSTSTOCKIST.MoveNext
        i = i + 1
    Loop
    RSTSTOCKIST.Close
    Set RSTSTOCKIST = Nothing
    
    Exit Sub
    
ERRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        'If REPFLAG = False Then RSTREP.Close
        'If TMPFLAG = False Then rstTMP.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.Height = 15555
    End If
   Cancel = CLOSEALL
End Sub
