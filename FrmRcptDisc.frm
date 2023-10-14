VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FRMRcptDisc 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dicounts & Other Credit Notes"
   ClientHeight    =   2775
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7080
   Icon            =   "FrmRcptDisc.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2775
   ScaleWidth      =   7080
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   4065
      TabIndex        =   10
      Top             =   1890
      Width           =   2955
      Begin VB.CommandButton cmdSAVE 
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
         Height          =   450
         Left            =   90
         TabIndex        =   0
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   1590
         TabIndex        =   1
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00E0E0E0&
      Height          =   2775
      Left            =   0
      TabIndex        =   3
      Top             =   -15
      Width           =   7050
      Begin VB.TextBox TXTREFNO 
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   3810
         MaxLength       =   20
         TabIndex        =   13
         Top             =   1395
         Width           =   3165
      End
      Begin VB.TextBox txtBillNo 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   420
         TabIndex        =   11
         Top             =   765
         Width           =   795
      End
      Begin VB.TextBox txtrcptamt 
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   1020
         MaxLength       =   8
         TabIndex        =   2
         Top             =   1425
         Width           =   1770
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5340
         TabIndex        =   15
         Top             =   765
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711935
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   " "
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1140
         TabIndex        =   17
         Top             =   195
         Width           =   4935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   105
         TabIndex        =   16
         Top             =   225
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   8
         Left            =   2925
         TabIndex        =   14
         Top             =   1425
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Pymnt"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   1
         Left            =   3870
         TabIndex        =   8
         Top             =   765
         Width           =   1395
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   0
         Left            =   105
         TabIndex        =   7
         Top             =   765
         Width           =   570
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   105
         TabIndex        =   6
         Top             =   1425
         Width           =   915
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   1245
         TabIndex        =   5
         Top             =   765
         Width           =   1350
      End
      Begin VB.Label LBLDATE 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   360
         Left            =   2580
         TabIndex        =   4
         Top             =   765
         Width           =   1215
      End
   End
   Begin VB.Label lblactcode 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblactcode"
      Height          =   315
      Left            =   1065
      TabIndex        =   12
      Top             =   3210
      Width           =   1620
   End
   Begin VB.Label lbltmprcptamt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "rcpt amount"
      Height          =   315
      Left            =   3150
      TabIndex        =   9
      Top             =   3285
      Width           =   1620
   End
End
Attribute VB_Name = "FRMRcptDisc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean

Private Sub cmdcancel_Click()
    'CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Date", vbOKOnly, "Receipt..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Receipt..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    
    If Val(txtrcptamt.Text) = 0 Then
        MsgBox "Enter Payment Amount", vbOKOnly, "Receipt..."
        txtrcptamt.SetFocus
        Exit Sub
    End If

    
    On Error GoTo Errhand
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "CN"
    RSTTRXFILE!INV_TRX_TYPE = "CN"
    RSTTRXFILE!CR_NO = Val(txtBillNo.Text)
    'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
    RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!RCPT_AMT = Val(txtrcptamt.Text)
    RSTTRXFILE!ACT_CODE = lblactcode.Caption
    RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
    RSTTRXFILE!INV_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.Text)
    RSTTRXFILE!INV_AMT = Null
    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE!BANK_FLAG = "N"
    RSTTRXFILE!B_TRX_TYPE = Null
    'RSTTRXFILE!B_TRX_NO = Null
    RSTTRXFILE!B_BILL_TRX_TYPE = Null
    RSTTRXFILE!B_TRX_YEAR = Null
    RSTTRXFILE!BANK_CODE = Null
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    MsgBox "Saved Successfully....", vbOKOnly, "RECEIPTS"
    Unload Me
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtrcptamt.SetFocus
    End Select
End Sub


Private Sub Form_Activate()
      txtrcptamt.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo Errhand
    
    AGNT_FLAG = True
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'CN'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
     
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    'Width = 8900
    'Height = 4485
    Left = 800
    Top = 1000
    
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
    FRMRcptReg.Enabled = True
    FRMRcptReg.GRDTranx.SetFocus
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "  /  /    " Then
                TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
                txtrcptamt.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                txtrcptamt.SetFocus
            End If
    End Select
End Sub

Private Sub txtrcptamt_GotFocus()
    txtrcptamt.SelStart = 0
    txtrcptamt.SelLength = Len(txtrcptamt.Text)
End Sub

Private Sub txtrcptamt_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(txtrcptamt.Text) = 0 Then
                MsgBox "Enter Payment Amount", vbOKOnly, "PAYMENT..."
                txtrcptamt.SetFocus
                Exit Sub
            End If
            TXTREFNO.SetFocus
        Case vbKeyEscape
            Unload Me
    End Select
End Sub

Private Sub txtrcptamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREFNO_GotFocus()
    TXTREFNO.SelStart = 0
    TXTREFNO.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTREFNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CmdSave.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub
