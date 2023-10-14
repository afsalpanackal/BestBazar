VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmcompindex 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   5820
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   10320
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "CompIndex.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "CompIndex.frx":030A
   ScaleHeight     =   102.658
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   182.034
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   -255
      TabIndex        =   6
      Top             =   -450
      Width           =   11850
      Begin VB.Label LBLSHOP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MEDICALS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   450
         TabIndex        =   7
         Top             =   120
         Width           =   11010
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   5160
      Left            =   0
      TabIndex        =   1
      Top             =   -30
      Width           =   9840
      Begin VB.Frame frmecompany 
         Height          =   1740
         Left            =   75
         TabIndex        =   14
         Top             =   2940
         Width           =   4440
         Begin VB.CommandButton cmdcancel 
            BackColor       =   &H00400000&
            Caption         =   "&CANCEL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   3375
            MaskColor       =   &H80000007&
            TabIndex        =   20
            Top             =   1065
            UseMaskColor    =   -1  'True
            Width           =   915
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00400000&
            Caption         =   "&SAVE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2235
            MaskColor       =   &H80000007&
            TabIndex        =   19
            Top             =   1080
            UseMaskColor    =   -1  'True
            Width           =   915
         End
         Begin VB.TextBox TXTITEM 
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
            Height          =   300
            Left            =   1170
            TabIndex        =   17
            Top             =   630
            Width           =   3105
         End
         Begin VB.TextBox TxtItemcode 
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
            Height          =   300
            Left            =   1620
            TabIndex        =   15
            Top             =   180
            Width           =   1455
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   1
            Left            =   75
            TabIndex        =   18
            Top             =   630
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY CODE"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   60
            TabIndex        =   16
            Top             =   195
            Width           =   1530
         End
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
         Height          =   4980
         Left            =   6840
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   120
         Width           =   2880
      End
      Begin VB.TextBox tXTMEDICINE 
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
         Height          =   345
         Left            =   0
         TabIndex        =   0
         Top             =   315
         Width           =   3645
      End
      Begin MSDataListLib.DataList LSTDISTI 
         Height          =   1815
         Left            =   3675
         TabIndex        =   2
         Top             =   120
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   3201
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
      Begin VB.ListBox LSTDUMMY 
         Height          =   1425
         Left            =   6975
         TabIndex        =   5
         Top             =   3360
         Visible         =   0   'False
         Width           =   2580
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1230
         Left            =   15
         TabIndex        =   10
         Top             =   675
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   2170
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
      Begin VB.Frame frmecontrol 
         BackColor       =   &H00C0E0FF&
         Height          =   735
         Left            =   645
         TabIndex        =   4
         Top             =   1890
         Width           =   5655
         Begin VB.CommandButton cmditemcreate 
            Caption         =   "&Create company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   135
            TabIndex        =   13
            Top             =   165
            Width           =   1215
         End
         Begin VB.CommandButton CMDREMOVE 
            BackColor       =   &H00400000&
            Caption         =   "RE&MOVE DISTRIBUTOR"
            Height          =   465
            Left            =   4230
            MaskColor       =   &H80000007&
            TabIndex        =   12
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   1335
         End
         Begin VB.CommandButton CMDADDDIST 
            BackColor       =   &H00400000&
            Caption         =   "ADD DIS&TRIBUTOR"
            Height          =   465
            Left            =   2730
            MaskColor       =   &H80000007&
            TabIndex        =   11
            Top             =   180
            UseMaskColor    =   -1  'True
            Width           =   1305
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
            Height          =   480
            Left            =   1440
            TabIndex        =   3
            Top             =   180
            Width           =   1185
         End
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "COMPANY"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   30
         TabIndex        =   9
         Top             =   105
         Width           =   3570
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmcompindex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim rstTMP As New ADODB.Recordset

Dim PHYFLAG As Boolean 'PHY
Dim TMPFLAG As Boolean 'TMP
Dim REPFLAG As Boolean 'REP

Dim M_EDIT As Boolean
Dim M_FLAG As Boolean

Dim k As Integer
Dim CLOSEALL As Integer


Private Sub CMDADDDIST_Click()
    
    Dim RSTA As ADODB.Recordset
    
    If DataList2.BoundText = "" Then
        MsgBox "SELECT THE ITEM", vbOKOnly, "ORDER"
        tXTMEDICINE.SetFocus
        Exit Sub
    End If
    
    If lstmanufact.SelCount = 0 Then
        MsgBox "Please Select the Distributor to be added", vbOKOnly, "ORDER"
        Exit Sub
    End If
    
    On Error GoTo eRRhAND
    
    i = 0
    
    Set RSTA = New ADODB.Recordset
    RSTA.Open "SELECT *  FROM COMPINDEX WHERE COMP_CODE = '" & DataList2.BoundText & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    With RSTA
        For i = 0 To lstmanufact.ListCount - 1
            If lstmanufact.Selected(i) Then
                .AddNew
                !ACT_CODE = Mid(lstmanufact.List(i), 1, 6)
                !COMP_CODE = DataList2.BoundText
                !COMP_NAME = DataList2.Text
                .Update
            End If
        Next i

        .Close
        
    End With
    
    
    Set RSTA = Nothing

    DataList2_Click
    LSTDISTI.SetFocus
       
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub cmdcancel_Click()
    Dim rstTMP As ADODB.Recordset
    
    TXTITEMCODE.Enabled = True
    TXTITEM.Text = ""
    TXTITEM.Enabled = False
    cmdSAVE.Enabled = False
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select MAX(Val(COMP_CODE)) From COMPMAST", db2, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        If IsNull(rstTMP.Fields(0)) Then
            TXTITEMCODE.Text = 1000
        Else
            TXTITEMCODE.Text = Val(rstTMP.Fields(0)) + 1
        End If
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    TXTITEMCODE.SetFocus
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDREMOVE_Click()
       
    If DataList2.BoundText = "" Then
        Exit Sub
    End If
    
    If LSTDISTI.Text = "" Then
        Exit Sub
    End If

    If MsgBox("ARE YOU SURE YOU WANT TO REMOVE " & LSTDISTI.Text, vbYesNo, "DELETING....") = vbNo Then Exit Sub
    On Error GoTo eRRhAND
      
    db2.Execute ("DELETE *  FROM COMPINDEX WHERE COMP_CODE = '" & DataList2.BoundText & "' AND ACT_CODE = '" & LSTDISTI.BoundText & "'")
    DataList2_Click
    LSTDISTI.SetFocus
       
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub


Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdSAVE.Enabled = False
            TXTITEM.Enabled = True
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub DataList2_Click()
    
    Dim RSTAVL As ADODB.Recordset
    Dim RSTMAN As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    Dim i As Integer
    Dim N As Integer
    
    On Error GoTo eRRhAND
    
    LSTDUMMY.Clear
    lstmanufact.Clear
    If TMPFLAG = True Then
        rstTMP.Open "SELECT * FROM ACTMAST RIGHT JOIN COMPINDEX ON ACTMAST.ACT_CODE = COMPINDEX.ACT_CODE WHERE COMP_CODE = '" & DataList2.BoundText & "' ORDER BY ACTMAST.ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        TMPFLAG = False
    Else
        rstTMP.Close
        rstTMP.Open "SELECT * FROM ACTMAST RIGHT JOIN COMPINDEX ON ACTMAST.ACT_CODE = COMPINDEX.ACT_CODE WHERE COMP_CODE = '" & DataList2.BoundText & "' ORDER BY ACTMAST.ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
        TMPFLAG = False
    End If
            
           
    Set Me.LSTDISTI.RowSource = rstTMP
    LSTDISTI.ListField = "ACT_NAME"
    LSTDISTI.BoundColumn = "ACT_CODE"
    
    
            
    i = 0
    Do Until rstTMP.EOF
        LSTDUMMY.AddItem (i)
        LSTDUMMY.List(i) = rstTMP!ACT_CODE
        i = i + 1
        rstTMP.MoveNext
    Loop
    
    i = 0
    Set RSTMAN = New ADODB.Recordset
    RSTMAN.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
    With RSTMAN
        Do Until .EOF
            
            For N = 0 To LSTDUMMY.ListCount
                If Trim(LSTDUMMY.List(N)) = Trim(!ACT_CODE) Then GoTo SKIP
            Next N
            lstmanufact.AddItem (i)
            lstmanufact.List(i) = !ACT_CODE & " " & Trim(!ACT_NAME)
            i = i + 1
SKIP:
        .MoveNext
        Loop
        .Close
    End With
    Set RSTMAN = Nothing

    Exit Sub
    
eRRhAND:
    MsgBox Err.Description

    
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRhAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM COMPMAST WHERE COMP_CODE = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTITEM.Text = RSTITEMMAST!COMP_NAME
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            TXTITEMCODE.Enabled = False
            TXTITEM.Enabled = True
            
            TXTITEM.SetFocus
        
    End Select
Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then
                MsgBox "ENTER COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            TXTITEM.Enabled = False
            cmdSAVE.Enabled = True
            cmdSAVE.SetFocus
            
        Case vbKeyEscape
            cmdSAVE.Enabled = False
            TXTITEM.Enabled = False
            TXTITEMCODE.Enabled = True
            TXTITEMCODE.SetFocus
    End Select
    
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    Dim rstTMP As ADODB.Recordset
    
    On Error GoTo eRRhAND
    PHYFLAG = True
    TMPFLAG = True
    REPFLAG = True

    CLOSEALL = 1
    
    TXTITEM.Enabled = False
    cmdSAVE.Enabled = False
    
    Me.Width = 10000
    Me.Height = 5800
    Me.Left = 0
    Me.Top = 0
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select MAX(Val(COMP_CODE)) From COMPMAST", db2, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        If IsNull(rstTMP.Fields(0)) Then
            TXTITEMCODE.Text = 1000
        Else
            TXTITEMCODE.Text = Val(rstTMP.Fields(0)) + 1
        End If
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If TMPFLAG = False Then rstTMP.Close
        If REPFLAG = False Then RSTREP.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
   Cancel = CLOSEALL
End Sub

Private Sub lstmanufact_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyReturn
            CMDADDDIST.SetFocus
                        
    End Select
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo eRRhAND
    If REPFLAG = True Then
        RSTREP.Open "Select * From COMPMAST  WHERE COMP_NAME Like '" & Me.tXTMEDICINE.Text & "%'ORDER BY [COMP_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select * From COMPMAST  WHERE COMP_NAME Like '" & Me.tXTMEDICINE.Text & "%'ORDER BY [COMP_NAME]", db2, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "COMP_NAME"
    DataList2.BoundColumn = "COMP_CODE"

    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then
                TXTUNIT.SetFocus
                Exit Sub
            End If
            CMDADDDIST.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstTMP As ADODB.Recordset
    
    If TXTITEM.Text = "" Then
        MsgBox "ENTER NAME OF COMPANY", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRhAND
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM COMPMAST WHERE COMP_CODE = '" & TXTITEMCODE.Text & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!COMP_NAME = TXTITEM.Text
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!COMP_CODE = TXTITEMCODE.Text
        RSTITEMMAST!COMP_NAME = TXTITEM.Text
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    tXTMEDICINE.Text = TXTITEM.Text
    DataList2.BoundText = TXTITEMCODE.Text
    Call DataList2_Click
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select MAX(Val(COMP_CODE)) From COMPMAST", db2, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        If IsNull(rstTMP.Fields(0)) Then
            TXTITEMCODE.Text = 1000
        Else
            TXTITEMCODE.Text = Val(rstTMP.Fields(0)) + 1
        End If
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    cmdcancel_Click
    
Exit Sub
eRRhAND:
    If Err.Number = -2147217873 Then
        MsgBox "ITEM ALREADY EXISTS", vbOKOnly, "ITEM CREATION..."
        cmdSAVE.Enabled = False
        TXTITEM.Enabled = True
        TXTITEM.SetFocus
    Else
        MsgBox (Err.Description)
    End If
        
End Sub
