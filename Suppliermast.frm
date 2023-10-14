VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmsuppliermast 
   BackColor       =   &H00C7F3DE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplier Creation"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   Icon            =   "Suppliermast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   7815
   Begin VB.TextBox txtsupplist 
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
      Left            =   1680
      MaxLength       =   34
      TabIndex        =   1
      Top             =   465
      Visible         =   0   'False
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      BackColor       =   &H00B7F0D5&
      Height          =   7275
      Left            =   120
      TabIndex        =   16
      Top             =   1245
      Width           =   7650
      Begin VB.TextBox txtML 
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
         Left            =   4425
         MaxLength       =   25
         TabIndex        =   36
         Top             =   5085
         Width           =   2235
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Import Suppliers"
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   6690
         Width           =   1305
      End
      Begin VB.CheckBox chkIGST 
         BackColor       =   &H00800080&
         Caption         =   "&IGST"
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
         Left            =   4515
         TabIndex        =   34
         Top             =   2385
         Width           =   1980
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
         Left            =   1590
         MaxLength       =   20
         TabIndex        =   31
         Top             =   1515
         Width           =   2895
      End
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
         Left            =   4530
         TabIndex        =   30
         Top             =   1890
         Width           =   1965
      End
      Begin VB.TextBox Txtopbal 
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
         Left            =   1590
         MaxLength       =   12
         TabIndex        =   28
         Top             =   5520
         Width           =   2235
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
         Left            =   120
         MaskColor       =   &H80000007&
         TabIndex        =   27
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   1260
      End
      Begin VB.TextBox txtcst 
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
         Left            =   1590
         MaxLength       =   25
         TabIndex        =   11
         Top             =   5085
         Width           =   2235
      End
      Begin VB.TextBox txtkgst 
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   10
         Top             =   4710
         Width           =   2235
      End
      Begin VB.TextBox txtremarks 
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
         Left            =   1590
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4320
         Width           =   6000
      End
      Begin VB.TextBox txtdlno 
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
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   8
         Top             =   3945
         Width           =   6000
      End
      Begin VB.TextBox txtemail 
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
         Left            =   1590
         MaxLength       =   150
         TabIndex        =   7
         Top             =   3540
         Width           =   6000
      End
      Begin VB.TextBox txtfaxno 
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
         Left            =   1590
         MaxLength       =   15
         TabIndex        =   6
         Top             =   3165
         Width           =   2880
      End
      Begin VB.TextBox txttelno 
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
         Left            =   1590
         MaxLength       =   25
         TabIndex        =   5
         Top             =   2790
         Width           =   2880
      End
      Begin VB.TextBox txtsupplier 
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
         Left            =   1590
         MaxLength       =   100
         TabIndex        =   3
         Top             =   225
         Width           =   6000
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
         Left            =   3510
         MaskColor       =   &H80000007&
         TabIndex        =   12
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Left            =   4545
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   915
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
         Left            =   5580
         MaskColor       =   &H80000007&
         TabIndex        =   14
         Top             =   6075
         UseMaskColor    =   -1  'True
         Width           =   915
      End
      Begin MSDataListLib.DataList Datacompany 
         Height          =   780
         Left            =   1590
         TabIndex        =   32
         Top             =   1890
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
         Caption         =   "ML No."
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
         Index           =   13
         Left            =   3870
         TabIndex        =   37
         Top             =   5115
         Width           =   1035
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
         Index           =   7
         Left            =   75
         TabIndex        =   33
         Top             =   1710
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "OP. Balance"
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
         Left            =   150
         TabIndex        =   29
         Top             =   5550
         Width           =   1350
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DL No."
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
         Top             =   5115
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "GSTIN No."
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
         Top             =   4740
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "IFS Code"
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
         Top             =   4350
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A/C No"
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
         Top             =   3975
         Width           =   1155
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
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
         Top             =   3570
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Mob No."
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
         Top             =   3195
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
         Top             =   2820
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
         Height          =   855
         Left            =   1590
         TabIndex        =   4
         Top             =   570
         Width           =   6000
         VariousPropertyBits=   -1400879077
         ForeColor       =   255
         MaxLength       =   150
         BorderStyle     =   1
         Size            =   "10583;1508"
         BorderColor     =   0
         SpecialEffect   =   0
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Name"
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
         Left            =   150
         TabIndex        =   17
         Top             =   255
         Width           =   1365
      End
   End
   Begin VB.TextBox Txtsuplcode 
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
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   465
      Width           =   1455
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1680
      TabIndex        =   2
      Top             =   825
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2858
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
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
      Caption         =   "Supplier Code"
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
Attribute VB_Name = "frmsuppliermast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
Dim COMPANYFLAG As Boolean
Dim REPFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset

Private Sub cmdcancel_Click()
    FRAME.Visible = False
    txtsupplier.text = ""
    txtaddress.text = ""
    txttelno.text = ""
    txtfaxno.text = ""
    txtemail.text = ""
    txtdlno.text = ""
    TXTREMARKS.text = ""
    txtkgst.text = ""
    TxtCST.text = ""
    txtML.text = ""
    Txtopbal.text = ""
    txtcompany.text = ""
    chknewcomp.Value = 0
    chkIGST.Value = 0
    Txtsuplcode.Enabled = True
End Sub

Private Sub CmdDelete_Click()
    Dim RSTSUPMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From RTRXFILE WHERE M_USER_ID = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRANSMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    Set RSTSUPMAST = New ADODB.Recordset
    RSTSUPMAST.Open "SELECT M_USER_ID From TRXMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
    If RSTSUPMAST.RecordCount > 0 Then
        MsgBox "CANNOT DELETE SINCE TRANSACTIONS FOUND!!!!", vbOKOnly, "DELETE!!!!"
        Exit Sub
    End If
    RSTSUPMAST.Close
    Set RSTSUPMAST = Nothing
    
    If (MsgBox("ARE YO SURE YOU WANT TO DELETE !!!!", vbYesNo, "SALES") = vbNo) Then Exit Sub
    db.Execute ("delete  FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    db.Execute ("delete  FROM PRODLINK WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'")
    Call cmdcancel_Click
    MsgBox "DELETED SUCCESSFULLY!!!!", vbOKOnly, "DELETE!!!!"
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    
    If txtsupplier.text = "" Then
        MsgBox "ENTER NAME OF SUPPLIER", vbOKOnly, "SUPPLIER MASTER"
        txtsupplier.SetFocus
        Exit Sub
    End If
    
    If chknewcomp.Value = 0 And Datacompany.BoundText = "" And txtcompany.text <> "" Then
        MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "Supplier Master"
        txtcompany.SetFocus
        Exit Sub
    End If
    
    If Trim(txtkgst.text) <> "" Then
        If Len(Trim(txtkgst.text)) <> 15 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SUPPLIER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If
        
        If Val(Left(Trim(txtkgst.text), 2)) = 0 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SUPPLIER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If
        
'        If Val(Mid(Trim(txtkgst.Text), 13, 1)) = 0 Then
'            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SUPPLIER MASTER"
'            txtkgst.SetFocus
'            Exit Sub
'        End If
        
        If Val(Mid(Trim(txtkgst.text), 14, 1)) <> 0 Then
            MsgBox "PLEASE ENTER A VALID GSTIN NO.", vbOKOnly, "SUPPLIER MASTER"
            txtkgst.SetFocus
            Exit Sub
        End If
    End If
    On Error GoTo ERRHAND
    Txtsuplcode.text = Format(Txtsuplcode.text, "000")
    Txtsuplcode.Tag = "311" & Txtsuplcode.text
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
    
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.text)
        RSTITEMMAST!Address = Trim(txtaddress.text)
        RSTITEMMAST!TELNO = Trim(txttelno.text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.text)
        RSTITEMMAST!DL_NO = Trim(txtdlno.text)
        RSTITEMMAST!REMARKS = Trim(TXTREMARKS.text)
        RSTITEMMAST!KGST = Trim(txtkgst.text)
        RSTITEMMAST!CST = Trim(TxtCST.text)
        RSTITEMMAST!ML_NO = Trim(txtML.text)
        
        RSTITEMMAST!OPEN_DB = Round(Val(Txtopbal.text), 3)
        If chkIGST.Value = 1 Then
            RSTITEMMAST!CUST_IGST = "Y"
        Else
            RSTITEMMAST!CUST_IGST = ""
        End If
        If txtcompany.text <> "" Or Datacompany.BoundText <> "" Then
            If chknewcomp.Value = 1 Then RSTITEMMAST!Area = txtcompany.text Else RSTITEMMAST!Area = Datacompany.BoundText
        End If
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!ACT_CODE = Txtsuplcode.Tag
        RSTITEMMAST!ACT_NAME = Trim(txtsupplier.text)
        RSTITEMMAST!Address = Trim(txtaddress.text)
        RSTITEMMAST!TELNO = Trim(txttelno.text)
        RSTITEMMAST!FAXNO = Trim(txtfaxno.text)
        RSTITEMMAST!EMAIL_ADD = Trim(txtemail.text)
        RSTITEMMAST!DL_NO = Trim(txtdlno.text)
        RSTITEMMAST!REMARKS = Trim(TXTREMARKS.text)
        RSTITEMMAST!KGST = Trim(txtkgst.text)
        RSTITEMMAST!CST = Trim(TxtCST.text)
        RSTITEMMAST!ML_NO = Trim(txtML.text)
        RSTITEMMAST!CONTACT_PERSON = "CS"
        RSTITEMMAST!SLSM_CODE = "SM"
        RSTITEMMAST!OPEN_DB = Round(Val(Txtopbal.text), 3)
        RSTITEMMAST!OPEN_CR = 0
        RSTITEMMAST!YTD_DB = 0
        RSTITEMMAST!YTD_CR = 0
        RSTITEMMAST!CUST_TYPE = ""
        RSTITEMMAST!CREATE_DATE = Date
        RSTITEMMAST!C_USER_ID = "SM"
        RSTITEMMAST!MODIFY_DATE = Date
        RSTITEMMAST!M_USER_ID = "SM"
        If chkIGST.Value = 1 Then
            RSTITEMMAST!CUST_IGST = "Y"
        Else
            RSTITEMMAST!CUST_IGST = ""
        End If
        If txtcompany.text <> "" Or Datacompany.BoundText <> "" Then
            If chknewcomp.Value = 1 Then RSTITEMMAST!Area = txtcompany.text Else RSTITEMMAST!Area = Datacompany.BoundText
        End If

        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    FRAME.Visible = False
    txtsupplier.text = ""
    txtaddress.text = ""
    txttelno.text = ""
    txtfaxno.text = ""
    txtemail.text = ""
    txtdlno.text = ""
    TXTREMARKS.text = ""
    txtkgst.text = ""
    TxtCST.text = ""
    txtML.text = ""
    Txtopbal.text = ""
    Txtsuplcode.Enabled = True
    txtcompany.text = ""
    chknewcomp.Value = 0
    
    On Error GoTo ERRHAND
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "Select MAX(ACT_CODE) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3)", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        Txtsuplcode.text = IIf(IsNull(RSTITEMMAST.Fields(0)), "001", Mid(RSTITEMMAST.Fields(0), 4) + 1)
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
Exit Sub
ERRHAND:
    MsgBox (err.Description)
        
End Sub

Private Sub Command2_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Import Customers"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Import Stock Items") = vbNo Then Exit Sub
    If MsgBox("Sheet Name should be 'SUPPLIERS' and First coloumn should be Supplier Name and Second coloumn should be Supplier Address", vbYesNo, "Import Suppliers") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Excel Files (*.xls*)|*.xls*"
    CommonDialog1.ShowOpen
    
    Screen.MousePointer = vbHourglass
    Set xlApp = New Excel.Application
    
    'Set wb = xlApp.Workbooks.Open("PATH TO YOUR EXCEL FILE")
    Set wb = xlApp.Workbooks.Open(CommonDialog1.FileName)
    
    Set ws = wb.Worksheets("SUPPLIERS") 'Specify your worksheet name
    var = ws.Range("A1").Value
    
'    db.Execute "dELETE FROM ACTMAST"
'    db.Execute "dELETE FROM RTRXFILE"
    
    Dim RstCustmast As ADODB.Recordset
    Dim RSTITEMTRX As ADODB.Recordset
    Dim CUSTCODE As String
    Dim sl As Integer
    Dim lastno As Integer
    sl = 1
    
    Set RstCustmast = New ADODB.Recordset
    RstCustmast.Open "Select MAX(act_code) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ", db, adOpenStatic, adLockReadOnly
    If Not (RstCustmast.EOF And RstCustmast.BOF) Then
        If IsNull(RstCustmast.Fields(0)) Then
            CUSTCODE = 1
        Else
            CUSTCODE = Val(RstCustmast.Fields(0)) + 1
        End If
    End If
    RstCustmast.Close
    Set RstCustmast = Nothing
        
    For i = 2 To 30000
        If Trim(ws.Range("A" & i).Value) = "" Then Exit For
        
        Set RstCustmast = New ADODB.Recordset
        RstCustmast.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & CUSTCODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        db.BeginTrans
        
        RstCustmast.AddNew
        'RSTCUSTMAST.Fields("PHOTO").AppendChunk bytData
        RstCustmast!ACT_CODE = "311" & Format(CUSTCODE, "000")
        RstCustmast!ACT_NAME = Trim(ws.Range("A" & i).Value)
        RstCustmast!Address = Trim(ws.Range("B" & i).Value)
        RstCustmast!TELNO = Trim(ws.Range("C" & i).Value)
        RstCustmast!FAXNO = Trim(ws.Range("D" & i).Value)
        RstCustmast!EMAIL_ADD = ""
        RstCustmast!DL_NO = ""
        RstCustmast!REMARKS = ""
        RstCustmast!KGST = Trim(ws.Range("E" & i).Value)
        RstCustmast!CST = ""
        RstCustmast!PYMT_PERIOD = 0
        RstCustmast!Area = Trim(ws.Range("F" & i).Value)
        'RstCustmast!AGENT_CODE = ""
        'RstCustmast!AGENT_NAME = ""
        'RstCustmast!Sl_no = CUSTCODE
        RstCustmast!CONTACT_PERSON = "CS"
        RstCustmast!SLSM_CODE = "SM"
        RstCustmast!OPEN_DB = Val(ws.Range("G" & i).Value)
        RstCustmast!OPEN_CR = 0
        RstCustmast!YTD_DB = 0
        RstCustmast!YTD_CR = 0
        RstCustmast!CREATE_DATE = Date
        RstCustmast!C_USER_ID = "SM"
        RstCustmast!MODIFY_DATE = Date
        RstCustmast!M_USER_ID = "SM"
        RstCustmast!Type = "W"
        RstCustmast!CUST_TYPE = ""
        RstCustmast!CUST_IGST = ""
        
        RstCustmast.Update
        RstCustmast.Close
        Set RstCustmast = Nothing
        db.CommitTrans
        CUSTCODE = CUSTCODE + 1
                        
SKIP:
    Next i
    wb.Close
    
    xlApp.Quit
    
    Set ws = Nothing
    Set wb = Nothing
    Set xlApp = Nothing
    Screen.MousePointer = vbNormal
        
    MsgBox "Success", vbOKOnly
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = 9 Then
        MsgBox "NO SUCH FILE PRESENT!!", vbOKOnly, "IMPORT ITEMS"
        wb.Close
        xlApp.Quit
        Set ws = Nothing
        Set wb = Nothing
        Set xlApp = Nothing
    ElseIf err.Number = 32755 Then
        
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ACT_CODE FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                Txtsuplcode.text = Mid(RSTITEMMAST!ACT_CODE, Len(RSTITEMMAST!ACT_CODE) - 2, 3)
            End If
            txtsupplist.Visible = False
            DataList2.Visible = False
            Txtsuplcode.SetFocus
        Case vbKeyEscape
            txtsupplist.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub Form_Activate()
    If Txtsuplcode.Enabled = True Then Txtsuplcode.SetFocus
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    
    REPFLAG = True
    COMPANYFLAG = True
    'TMPFLAG = True
    'Me.Width = 6930
    'Me.Height = 8265
    Me.Left = 3500
    Me.Top = 100
    FRAME.Visible = False
    'txtunit.Visible = False
    On Error GoTo ERRHAND
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(ACT_CODE) From ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3)", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        Txtsuplcode.text = IIf(IsNull(TRXMAST.Fields(0)), "001", Mid(TRXMAST.Fields(0), 4) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub txtaddress_GotFocus()
    txtaddress.SelStart = 0
    txtaddress.SelLength = Len(txtaddress.text)
End Sub

Private Sub txtaddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtcompany.SetFocus
    End Select
End Sub

Private Sub txtaddress_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txtcst_GotFocus()
    TxtCST.SelStart = 0
    TxtCST.SelLength = Len(TxtCST.text)
End Sub

Private Sub txtcst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtopbal.SetFocus
    End Select
End Sub

Private Sub txtML_GotFocus()
    txtML.SelStart = 0
    txtML.SelLength = Len(txtML.text)
End Sub

Private Sub txtML_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtopbal.SetFocus
    End Select
End Sub


Private Sub txtdlno_GotFocus()
    txtdlno.SelStart = 0
    txtdlno.SelLength = Len(txtdlno.text)
End Sub

Private Sub txtdlno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTREMARKS.SetFocus
    End Select
End Sub

Private Sub txtemail_GotFocus()
    txtemail.SelStart = 0
    txtemail.SelLength = Len(txtemail.text)
End Sub

Private Sub txtemail_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtdlno.SetFocus
    End Select
End Sub

Private Sub txtfaxno_GotFocus()
    txtfaxno.SelStart = 0
    txtfaxno.SelLength = Len(txtfaxno.text)
End Sub

Private Sub txtfaxno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtemail.SetFocus
    End Select
End Sub

Private Sub txtkgst_GotFocus()
    txtkgst.SelStart = 0
    txtkgst.SelLength = Len(txtkgst.text)
End Sub

Private Sub txtkgst_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtCST.SetFocus
    End Select
End Sub

Private Sub txtkgst_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtopbal_GotFocus()
    Txtopbal.SelStart = 0
    Txtopbal.SelLength = Len(Txtopbal.text)
End Sub

Private Sub Txtopbal_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
    End Select
End Sub

Private Sub Txtopbal_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("."), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtkgst.SetFocus
    End Select
End Sub

Private Sub txtsupplier_GotFocus()
    txtsupplier.SelStart = 0
    txtsupplier.SelLength = Len(txtsupplier.text)
   
End Sub

Private Sub txtsupplier_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtsupplier.text = "" Then
                MsgBox "ENTER NAME OF SUPPLIER", vbOKOnly, "Supplier Master"
                txtsupplier.SetFocus
                Exit Sub
            End If
         txtaddress.SetFocus
    End Select
    
End Sub

Private Sub txtsupplier_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub Txtsuplcode_GotFocus()
    Txtsuplcode.SelStart = 0
    Txtsuplcode.SelLength = Len(Txtsuplcode.text)
End Sub

Private Sub Txtsuplcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(Txtsuplcode.text) = "" Then Exit Sub
            If Val(Txtsuplcode.text) = 0 Then Exit Sub
            On Error GoTo ERRHAND
            Txtsuplcode.text = Format(Txtsuplcode.text, "000")
            Txtsuplcode.Tag = "311" & Txtsuplcode.text
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & Txtsuplcode.Tag & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtsupplier.text = RSTITEMMAST!ACT_NAME
                txtaddress.text = IIf(IsNull(RSTITEMMAST!Address), "", RSTITEMMAST!Address)
                txttelno.text = IIf(IsNull(RSTITEMMAST!TELNO), "", RSTITEMMAST!TELNO)
                txtfaxno.text = IIf(IsNull(RSTITEMMAST!FAXNO), "", RSTITEMMAST!FAXNO)
                txtemail.text = IIf(IsNull(RSTITEMMAST!EMAIL_ADD), "", RSTITEMMAST!EMAIL_ADD)
                txtdlno.text = IIf(IsNull(RSTITEMMAST!DL_NO), "", RSTITEMMAST!DL_NO)
                TXTREMARKS.text = IIf(IsNull(RSTITEMMAST!REMARKS), "", RSTITEMMAST!REMARKS)
                txtkgst.text = IIf(IsNull(RSTITEMMAST!KGST), "", RSTITEMMAST!KGST)
                TxtCST.text = IIf(IsNull(RSTITEMMAST!CST), "", RSTITEMMAST!CST)
                txtML.text = IIf(IsNull(RSTITEMMAST!ML_NO), "", RSTITEMMAST!ML_NO)
                Txtopbal.text = IIf(IsNull(RSTITEMMAST!OPEN_DB), 0, RSTITEMMAST!OPEN_DB)
                If RSTITEMMAST!CUST_IGST = "Y" Then
                    chkIGST.Value = 1
                Else
                    chkIGST.Value = 0
                End If
                txtcompany.text = IIf(IsNull(RSTITEMMAST!Area), "", RSTITEMMAST!Area)
                Datacompany.text = txtcompany.text
                Call Datacompany_Click
                CmdDelete.Enabled = True
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
            
            Txtsuplcode.Enabled = False
            FRAME.Visible = True
            txtsupplier.SetFocus
        Case 114
            txtsupplist.text = ""
            txtsupplist.Visible = True
            DataList2.Visible = True
            txtsupplist.SetFocus
        Case vbKeyEscape
            Call CmdExit_Click
    End Select
Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Txtsuplcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtsupplist_Change()
    On Error GoTo ERRHAND
    If REPFLAG = True Then
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.txtsupplist.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
        'RSTREP.Open "Select ACT_CODE,ACT_NAME From ACTMAST  WHERE ACT_NAME Like '" & Me.txtsupplist.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic,adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ACT_NAME"
    DataList2.BoundColumn = "ACT_CODE"
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtsupplist_GotFocus()
    txtsupplist.SelStart = 0
    txtsupplist.SelLength = Len(txtsupplist.text)
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
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub txtsupplist_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub txttelno_GotFocus()
    txttelno.SelStart = 0
    txttelno.SelLength = Len(txttelno.text)
End Sub

Private Sub txttelno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtfaxno.SetFocus
    End Select
End Sub

Private Sub txtcompany_Change()
    On Error GoTo ERRHAND
    
    Set Me.Datacompany.RowSource = Nothing
    If COMPANYFLAG = True Then
        RSTCOMPANY.Open "Select DISTINCT AREA From ACTMAST WHERE AREA Like '" & txtcompany.text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    Else
        RSTCOMPANY.Close
        RSTCOMPANY.Open "Select DISTINCT AREA From ACTMAST WHERE AREA Like '" & txtcompany.text & "%' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        COMPANYFLAG = False
    End If
    Set Me.Datacompany.RowSource = RSTCOMPANY
    Datacompany.ListField = "AREA"
    Datacompany.BoundColumn = "AREA"
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcompany_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.text)
End Sub

Private Sub txtcompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            '''''''If txtcompany.Text = "" Then Exit Sub
            Datacompany.SetFocus
        Case vbKeyEscape
            TxtCST.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo ERRHAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT AREA FROM ACTMAST WHERE ACT_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                txtcompany.text = RSTITEMMAST!Area
            Else
'                If txtcompany.Text = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "CUSTOMER MASTER"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
'                If chknewcomp.Value = 0 And Datacompany.BoundText = "" Then
'                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "Supplier Master"
'                    txtcompany.SetFocus
'                    Exit Sub
'                End If
                If chknewcomp.Value = 0 And Datacompany.BoundText = "" And txtcompany.text <> "" Then
                    MsgBox "SELECT AREA FOR CUSTOMER", vbOKOnly, "Supplier Master"
                    txtcompany.SetFocus
                    Exit Sub
                End If
            End If
            txttelno.SetFocus
        Case vbKeyEscape
            txtcompany.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub Datacompany_Click()
'    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ERRHAND
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT MANUFACTURER FROM ITEMMAST WHERE ITEM_CODE = '" & Datacompany.BoundText & "'", db, adOpenStatic, adLockReadOnly
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        txtcompany.Text = RSTITEMMAST!MANUFACTURER
'    End If
    txtcompany.text = Datacompany.BoundText
    Datacompany.text = txtcompany.text
    Exit Sub
ERRHAND:
    MsgBox (err.Description)
End Sub

Private Sub chknewcomp_Click()
    On Error Resume Next
    txtcompany.SetFocus
End Sub
