VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmcustTRXFILE 
   BackColor       =   &H00B7DFCB&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Branch Office Creation"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   ControlBox      =   0   'False
   Icon            =   "CustTRXFILE.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6840
   Begin VB.TextBox txtsupplist 
      Appearance      =   0  'Flat
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
      Height          =   360
      Left            =   1815
      MaxLength       =   60
      TabIndex        =   1
      Top             =   450
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00B7DFCB&
      Height          =   5115
      Left            =   0
      TabIndex        =   16
      Top             =   840
      Width           =   6795
      Begin VB.CheckBox chknewcomp 
         BackColor       =   &H00800080&
         Caption         =   "&New Area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   255
         Left            =   4155
         TabIndex        =   29
         Top             =   8265
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtcompany 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   28
         Top             =   7890
         Visible         =   0   'False
         Width           =   2895
      End
      Begin VB.CommandButton cmddelete 
         BackColor       =   &H00400000&
         Caption         =   "&DELETE"
         Enabled         =   0   'False
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
         Left            =   4470
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   4515
         UseMaskColor    =   -1  'True
         Width           =   1100
      End
      Begin VB.TextBox txtcst 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   25
         TabIndex        =   11
         Top             =   3870
         Width           =   2235
      End
      Begin VB.TextBox txtkgst 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   25
         TabIndex        =   10
         Top             =   3495
         Width           =   2235
      End
      Begin VB.TextBox txtremarks 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   9
         Top             =   3105
         Width           =   4050
      End
      Begin VB.TextBox txtdlno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   40
         TabIndex        =   8
         Top             =   2730
         Width           =   4050
      End
      Begin VB.TextBox txtemail 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   50
         TabIndex        =   7
         Top             =   2325
         Width           =   4665
      End
      Begin VB.TextBox txtfaxno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   6
         Top             =   1950
         Width           =   2235
      End
      Begin VB.TextBox txttelno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   15
         TabIndex        =   5
         Top             =   1575
         Width           =   2235
      End
      Begin VB.TextBox txtsupplier 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         Left            =   1665
         MaxLength       =   60
         TabIndex        =   3
         Top             =   255
         Width           =   4980
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
         Left            =   2190
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   4515
         UseMaskColor    =   -1  'True
         Width           =   1100
      End
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
         Left            =   3330
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   4515
         UseMaskColor    =   -1  'True
         Width           =   1100
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00400000&
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
         Height          =   390
         Left            =   5610
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   4515
         UseMaskColor    =   -1  'True
         Width           =   1100
      End
      Begin MSDataListLib.DataList Datacompany 
         Height          =   780
         Left            =   1215
         TabIndex        =   30
         Top             =   8280
         Visible         =   0   'False
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1376
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16512
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Area"
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
         Index           =   6
         Left            =   -300
         TabIndex        =   31
         Top             =   8100
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "CST No."
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
         Index           =   5
         Left            =   150
         TabIndex        =   26
         Top             =   3900
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "TIN No."
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
         Index           =   4
         Left            =   150
         TabIndex        =   25
         Top             =   3525
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KGST NO.2"
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
         Index           =   3
         Left            =   150
         TabIndex        =   24
         Top             =   3135
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "KGST NO.1"
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
         Index           =   2
         Left            =   150
         TabIndex        =   23
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Email Address"
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
         Index           =   12
         Left            =   150
         TabIndex        =   22
         Top             =   2355
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mobile No."
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
         Index           =   11
         Left            =   150
         TabIndex        =   21
         Top             =   1980
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Telephone No."
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
         Index           =   10
         Left            =   150
         TabIndex        =   20
         Top             =   1605
         Width           =   1500
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
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
         Index           =   9
         Left            =   150
         TabIndex        =   19
         Top             =   645
         Width           =   1290
      End
      Begin MSForms.TextBox txtaddress 
         Height          =   870
         Left            =   1665
         TabIndex        =   4
         Top             =   630
         Width           =   4980
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   99
         BorderStyle     =   1
         Size            =   "8784;1535"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Branch Name"
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
         Left            =   135
         TabIndex        =   17
         Top             =   255
         Width           =   1515
      End
   End
   Begin VB.TextBox Txtsuplcode 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   360
      Left            =   1815
      MaxLength       =   15
      TabIndex        =   0
      Top             =   450
      Width           =   3045
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1500
      Left            =   1815
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2646
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLCUSTCODE 
      Height          =   330
      Left            =   5145
      TabIndex        =   32
      Top             =   60
      Width           =   1605
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 to Search...... Esc to exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   300
      Index           =   8
      Left            =   135
      TabIndex        =   18
      Top             =   45
      Width           =   6300
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Branch Code"
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
      Left            =   105
      TabIndex        =   15
      Top             =   465
      Width           =   1560
   End
End
Attribute VB_Name = "frmcustTRXFILE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim CLOSEALL As Integer
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset

Private Sub cmdcancel_Click()
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtremarks.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo ErrHand
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From RTRXFILE WHERE [M_USER_ID] = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRANSMAST WHERE [ACT_CODE] = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXMAST WHERE [ACT_CODE] = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "SALES") = vbNo) Then Exit Sub
    db.Execute ("DELETE  FROM CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' AND BR_CODE = '" & Txtsuplcode.Text & "'")
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.Text = "" Then
        MsgBox "ENTER NAME OF CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If
    
'    If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'        MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'        txtcompany.SetFocus
'        Exit Sub
'    End If
    

    On Error GoTo ErrHand
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' AND BR_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!BR_CODE = Txtsuplcode.Text
    End If
    RSTITEMMAST!ACT_CODE = LBLCUSTCODE.Caption
    RSTITEMMAST!BR_NAME = Trim(txtsupplier.Text)
    RSTITEMMAST!Address = Trim(txtaddress.Text)
    RSTITEMMAST!TELNO = Trim(txttelno.Text)
    RSTITEMMAST!FAXNO = Trim(txtfaxno.Text)
    RSTITEMMAST!EMAIL_ADD = Trim(txtemail.Text)
    RSTITEMMAST!DL_NO = Trim(txtdlno.Text)
    RSTITEMMAST!Remarks = Trim(txtremarks.Text)
    RSTITEMMAST!KGST = Trim(txtkgst.Text)
    RSTITEMMAST!CST = Trim(txtcst.Text)
    'RSTITEMMAST!CONTACT_PERSON = "CS"
    'RSTITEMMAST!SLSM_CODE = "SM"
    'RSTITEMMAST!OPEN_DB = 0
    'RSTITEMMAST!OPEN_CR = 0
    'RSTITEMMAST!YTD_DB = 0
    'RSTITEMMAST!YTD_CR = 0
    'RSTITEMMAST!CUST_TYPE = ""
    'RSTITEMMAST!CREATE_DATE = Date
    'RSTITEMMAST!C_USER_ID = "SM"
    'RSTITEMMAST!MODIFY_DATE = Date
    'RSTITEMMAST!M_USER_ID = "SM"
    'RSTITEMMAST!Type = "W"
    'RSTITEMMAST!Sl_no = Val(Txtsuplcode.Text)
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    FRAME.Visible = False
    txtsupplier.Text = ""
    txtaddress.Text = ""
    txttelno.Text = ""
    txtfaxno.Text = ""
    txtemail.Text = ""
    txtdlno.Text = ""
    txtremarks.Text = ""
    txtkgst.Text = ""
    txtcst.Text = ""
    txtcompany.Text = ""
    Txtsuplcode.Enabled = True
    frmcustmast1.cmdexit.Enabled = False
    frmcustmast1.cmdcancel.Enabled = False
Exit Sub
ErrHand:
    MsgBox (Err.Description)
        
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT [ACT_CODE] FROM CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' AND BR_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.Text = RSTITEMMAST!BR_CODE
            End If
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    'If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    CLOSEALL = 1
    'Me.Width = 7000
    'Me.Height = 8625
    Me.Left = 2500
    Me.Top = 0
    FRAME.Visible = False
    'txtunit.Visible = False
    LBLCUSTCODE.Caption = frmcustmast1.Txtsuplcode.Text
    On Error GoTo ErrHand
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(BR_CODE) From CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = IIf(IsNull(TRXMAST.Fields(0)), "1", TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    On Error GoTo ErrHand
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select * From CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' AND BR_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.Text = TRXMAST!ACT_CODE
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    Exit Sub
ErrHand:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If REPFLAG = False Then RSTREP.Close
        If COMPANYFLAG = False Then RSTCOMPANY.Close
        MDIMAIN.PCTMENU.Enabled = True
        frmcustmast1.Enabled = True
        frmcustmast1.SetFocus
        
        'MDIMAIN.PCTMENU.SetFocus
    End If
   Cancel = CLOSEALL
End Sub

Private Sub txtaddress_GotFocus()
    txtaddress.SelStart = 0
    txtaddress.SelLength = Len(txtaddress.Text)
End Sub

Private Sub txtaddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'txttelno.SetFocus
    End Select
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtcst_GotFocus()
    txtcst.SelStart = 0
    txtcst.SelLength = Len(txtcst.Text)
End Sub

Private Sub txtcst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
    End Select
End Sub

Private Sub txtdlno_GotFocus()
    txtdlno.SelStart = 0
    txtdlno.SelLength = Len(txtdlno.Text)
End Sub

Private Sub txtdlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtremarks.SetFocus
    End Select
End Sub

Private Sub txtemail_GotFocus()
    txtemail.SelStart = 0
    txtemail.SelLength = Len(txtemail.Text)
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtdlno.SetFocus
    End Select
End Sub

Private Sub txtfaxno_GotFocus()
    txtfaxno.SelStart = 0
    txtfaxno.SelLength = Len(txtfaxno.Text)
End Sub

Private Sub txtfaxno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtemail.SetFocus
    End Select
End Sub

Private Sub txtkgst_GotFocus()
    txtkgst.SelStart = 0
    txtkgst.SelLength = Len(txtkgst.Text)
End Sub

Private Sub txtkgst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcst.SetFocus
    End Select
End Sub

Private Sub txtremarks_GotFocus()
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtkgst.SetFocus
    End Select
End Sub

Private Sub txtsupplier_GotFocus()
    txtsupplier.SelStart = 0
    txtsupplier.SelLength = Len(txtsupplier.Text)
   
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtsupplier.Text = "" Then
                MsgBox "ENTER NAME OF CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
                txtsupplier.SetFocus
                Exit Sub
            End If
         txtaddress.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub Txtsuplcode_GotFocus()
    Txtsuplcode.SelStart = 0
    Txtsuplcode.SelLength = Len(Txtsuplcode.Text)
End Sub

Private Sub Txtsuplcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Txtsuplcode.Text) = "" Then Exit Sub
            'If Val(Txtsuplcode.Text) = 0 Then Exit Sub
            If Trim(Txtsuplcode.Text) = "130000" Then
                MsgBox "This Code Cannot be created!!!!", , "Customer Creation"
                Exit Sub
            End If
            
            On Error GoTo ErrHand
            Set RSTITEMMAST = New ADODB.Recordset
            'RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & Txtsuplcode.Text & "' and ACT_CODE <> '130000'", db, adOpenStatic, adLockReadOnly
            RSTITEMMAST.Open "SELECT * FROM CUSTTRXFILE WHERE ACT_CODE = '" & LBLCUSTCODE.Caption & "' AND BR_CODE = '" & Txtsuplcode.Text & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.Text = RSTITEMMAST!BR_NAME
                txtaddress.Text = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                txttelno.Text = IIf(IsNull(RSTITEMMAST!TELNO), "", RSTITEMMAST!TELNO)
                txtfaxno.Text = IIf(IsNull(RSTITEMMAST!FAXNO), "", RSTITEMMAST!FAXNO)
                txtemail.Text = IIf(IsNull(RSTITEMMAST!EMAIL_ADD), "", RSTITEMMAST!EMAIL_ADD)
                txtdlno.Text = IIf(IsNull(RSTITEMMAST!DL_NO), "", RSTITEMMAST!DL_NO)
                txtremarks.Text = IIf(IsNull(RSTITEMMAST!Remarks), "", RSTITEMMAST!Remarks)
                txtkgst.Text = IIf(IsNull(RSTITEMMAST!KGST), "", RSTITEMMAST!KGST)
                txtcst.Text = IIf(IsNull(RSTITEMMAST!CST), "", RSTITEMMAST!CST)
                cmddelete.Enabled = True
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Txtsuplcode.Enabled = False
            FRAME.Visible = True
            txtsupplier.SetFocus
        Case 114
            txtsupplist.Text = ""
            txtsupplist.Visible = True
            DataList2.Visible = True
            txtsupplist.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Txtsuplcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplist_Change()
    On Error GoTo ErrHand
    If REPFLAG = True Then
        RSTREP.Open "Select [ACT_CODE],[BR_NAME] From CUSTTRXFILE  WHERE BR_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select [ACT_CODE],[BR_NAME] From CUSTTRXFILE  WHERE BR_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY BR_NAME", db, adOpenStatic, adLockReadOnly
        'RSTREP.Open "Select [ACT_CODE],[ACT_NAME] From ACTMAST  WHERE ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY [ACT_NAME]", db, adOpenStatic,adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ACT_NAME"
    DataList2.BoundColumn = "ACT_CODE"
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub txtsupplist_GotFocus()
    txtsupplist.SelStart = 0
    txtsupplist.SelLength = Len(txtsupplist.Text)
End Sub

Private Sub txtsupplist_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If txtsupplist.Text = "" Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub txtsupplist_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txttelno_GotFocus()
    txttelno.SelStart = 0
    txttelno.SelLength = Len(txttelno.Text)
End Sub

Private Sub txttelno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtfaxno.SetFocus
    End Select
End Sub

