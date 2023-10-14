VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form FRMRECEIPT 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT..."
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   ControlBox      =   0   'False
   Icon            =   "FrmPayment.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   9915
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Height          =   780
      Left            =   5430
      TabIndex        =   29
      Top             =   2775
      Width           =   4305
      Begin VB.CommandButton cmdSAVE 
         Caption         =   "&Save"
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
         Height          =   450
         Left            =   90
         TabIndex        =   6
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "&Cancel"
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
         Height          =   450
         Left            =   1590
         TabIndex        =   7
         Top             =   210
         Width           =   1230
      End
      Begin VB.CommandButton CMDEXIT 
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
         Height          =   450
         Left            =   2940
         TabIndex        =   8
         Top             =   195
         Width           =   1260
      End
   End
   Begin VB.TextBox txtBillNo 
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
      Left            =   1500
      TabIndex        =   0
      Top             =   375
      Width           =   885
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00C0FFC0&
      Enabled         =   0   'False
      Height          =   3525
      Left            =   120
      TabIndex        =   9
      Top             =   150
      Width           =   9705
      Begin VB.TextBox TXTDEALER 
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
         Height          =   330
         Left            =   1395
         TabIndex        =   2
         Top             =   690
         Width           =   3735
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   105
         TabIndex        =   17
         Top             =   1725
         Width           =   9495
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bal. Amt"
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
            Left            =   7155
            TabIndex        =   25
            Top             =   360
            Width           =   810
         End
         Begin VB.Label lblbalamt 
            Alignment       =   2  'Center
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
            Left            =   8070
            TabIndex        =   24
            Top             =   300
            Width           =   1230
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Total Rcvd Amt"
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
            Height          =   465
            Index           =   12
            Left            =   4530
            TabIndex        =   23
            Top             =   315
            Width           =   1020
         End
         Begin VB.Label lblrcvdamt 
            Alignment       =   2  'Center
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
            Left            =   5715
            TabIndex        =   22
            Top             =   315
            Width           =   1155
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Dtd. "
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
            Left            =   75
            TabIndex        =   21
            Top             =   345
            Width           =   495
         End
         Begin VB.Label lblinvdate 
            Alignment       =   2  'Center
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
            Left            =   540
            TabIndex        =   20
            Top             =   315
            Width           =   1215
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Total Bill Amt"
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
            Left            =   1875
            TabIndex        =   19
            Top             =   375
            Width           =   1335
         End
         Begin VB.Label lblbillamt 
            Alignment       =   2  'Center
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
            Left            =   3285
            TabIndex        =   18
            Top             =   330
            Width           =   1155
         End
      End
      Begin VB.TextBox txtrcptamt 
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Left            =   6825
         MaxLength       =   8
         TabIndex        =   5
         Top             =   1260
         Width           =   1005
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   3690
         TabIndex        =   1
         Top             =   240
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
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
      Begin MSDataListLib.DataCombo cmbinv 
         Height          =   330
         Left            =   6810
         TabIndex        =   4
         Top             =   765
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ForeColor       =   255
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   645
         Left            =   1395
         TabIndex        =   3
         Top             =   1050
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   1138
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   105
         TabIndex        =   26
         Top             =   735
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Rcpt. Date"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2550
         TabIndex        =   16
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "PAYMENT NO"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   75
         TabIndex        =   15
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RECEIPT AMT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   5280
         TabIndex        =   14
         Top             =   1290
         Width           =   1335
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
         Height          =   300
         Index           =   3
         Left            =   5235
         TabIndex        =   13
         Top             =   255
         Width           =   1395
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
         Height          =   360
         Left            =   6795
         TabIndex        =   12
         Top             =   210
         Width           =   1215
      End
      Begin VB.Label LBLTIME 
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
         Height          =   360
         Left            =   8115
         TabIndex        =   11
         Top             =   210
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Against Bill No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   5280
         TabIndex        =   10
         Top             =   795
         Width           =   1455
      End
   End
   Begin VB.Label TMPENTRYDATE 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   0
      TabIndex        =   37
      Top             =   4920
      Width           =   1620
   End
   Begin VB.Label TMPPAIDAMT 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   8040
      TabIndex        =   36
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label TMPINVNO 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   6240
      TabIndex        =   35
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label TMPSUPPLIER 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   4560
      TabIndex        =   34
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label TMPACTCODE 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   2880
      TabIndex        =   33
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label TMPRCPTDATE 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   1200
      TabIndex        =   32
      Top             =   4440
      Width           =   1620
   End
   Begin VB.Label TMPRCPTNO 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   120
      TabIndex        =   31
      Top             =   4440
      Width           =   1020
   End
   Begin VB.Label LBLCOMBOLIST 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LBLCOMBOLIST"
      Height          =   315
      Left            =   3360
      TabIndex        =   30
      Top             =   3960
      Width           =   1620
   End
   Begin VB.Label lbldealer 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   600
      TabIndex        =   28
      Top             =   3960
      Width           =   1620
   End
   Begin VB.Label flagchange 
      BorderStyle     =   1  'Fixed Single
      Height          =   315
      Left            =   60
      TabIndex        =   27
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "FRMRECEIPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_FLAG As Boolean
Dim INV_FLAG As Boolean
Dim INV_REC As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim CLOSEALL As Integer

Private Sub cmbinv_Change()
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From CRDTPYMT WHERE INV_NO = " & Val(cmbinv.Text) & " AND TRX_TYPE = 'CR'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        LBLINVDATE.Caption = RSTTRXFILE!INV_DATE
        LBLBILLAMT.Caption = Format(RSTTRXFILE!INV_AMT, "0.00")
        lblrcvdamt.Caption = Format(RSTTRXFILE!RCPT_AMOUNT, "0.00")
        LBLBALAMT.Caption = Format(RSTTRXFILE!BAL_AMT, "0.00")
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmbinv_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If cmbinv.Text = "" Then
                MsgBox "Select the Invoice", vbOKOnly, "PAYMENT..."
                cmbinv.SetFocus
                Exit Sub
            End If
            If cmbinv.MatchedWithList = False Then
                MsgBox "Select the Invoice from list", vbOKOnly, "PAYMENT..."
                cmbinv.SetFocus
                Exit Sub
            End If
            
            cmbinv.Enabled = False
            txtrcptamt.Enabled = True
            txtrcptamt.SetFocus
        Case vbKeyEscape
            cmbinv.Enabled = False
            TXTDEALER.Enabled = True
            DataList2.Enabled = True
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub cmdcancel_Click()

    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRHAND
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from CRDTPYMT WHERE INV_NO = " & Val(TMPINVNO.Caption) & " AND TRX_TYPE='CR'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
        RSTTRXFILE!RCPT_AMOUNT = RSTTRXFILE!RCPT_AMOUNT + Val(TMPPAIDAMT.Caption)
        RSTTRXFILE!BAL_AMT = RSTTRXFILE!INV_AMT - RSTTRXFILE!RCPT_AMOUNT
        If RSTTRXFILE!BAL_AMT <= 0 Then RSTTRXFILE!CHECK_FLAG = "Y" Else RSTTRXFILE!CHECK_FLAG = "N"
        'RSTTRXFILE!CHECK_FLAG = "N"
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    If TMPRCPTNO.Caption <> "" Then
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRNXRCPT", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            RSTTRXFILE.AddNew
            RSTTRXFILE!RCPT_NO = Val(TMPRCPTNO.Caption)
            RSTTRXFILE!TRX_TYPE = "PY"
            RSTTRXFILE!INV_NO = Val(TMPINVNO.Caption)
            RSTTRXFILE!RCPT_DATE = Format(TMPRCPTDATE.Caption, "DD/MM/YYYY")
            RSTTRXFILE!RCPT_AMOUNT = Val(TMPPAIDAMT.Caption)
            RSTTRXFILE!ACT_CODE = TMPACTCODE.Caption
            RSTTRXFILE!ACT_NAME = TMPSUPPLIER.Caption
            RSTTRXFILE!RCPT_ENTRY_DATE = Format(TMPENTRYDATE.Caption, "DD/MM/YYYY")
            RSTTRXFILE.Update
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'PY'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    cmbinv.Text = ""
    txtrcptamt.Text = ""
    DataList2.Text = ""
    LBLDATE.Caption = Format(Date, "DD/MM/YYYY")
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    LBLINVDATE.Caption = ""
    LBLBILLAMT.Caption = ""
    lblrcvdamt.Caption = ""
    LBLBALAMT.Caption = ""
    TMPRCPTNO.Caption = ""
    TMPRCPTDATE.Caption = ""
    TMPACTCODE.Caption = ""
    TMPSUPPLIER.Caption = ""
    TMPINVNO.Caption = ""
    TMPPAIDAMT.Caption = ""
    TMPENTRYDATE.Caption = ""
    LBLCOMBOLIST.Caption = ""
    FRMEMAIN.Enabled = False
    txtBillNo.Enabled = True
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    Exit Sub
eRRHAND:
    MsgBox Err.Description

End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter Proper Date", vbOKOnly, "PAYMENT..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "PAYMENT..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    
    If IsNull(DataList2.SelectedItem) Then
        MsgBox "Select Supplier From List", vbOKOnly, "PAYMENT..."
        DataList2.SetFocus
        Exit Sub
    End If
    
    If Val(txtrcptamt.Text) = 0 Then
        MsgBox "Enter Payment Amount", vbOKOnly, "PAYMENT..."
        txtrcptamt.SetFocus
        Exit Sub
    End If
    
    If cmbinv.Text = "" Then
        MsgBox "Select the Invoice", vbOKOnly, "PAYMENT..."
        cmbinv.SetFocus
        Exit Sub
    End If
    On Error GoTo eRRHAND

    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TRNXRCPT WHERE TRX_TYPE='PY' AND RCPT_NO= " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMOUNT = Val(txtrcptamt.Text)
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!RCPT_ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
        RSTTRXFILE.Update
    Else
        RSTTRXFILE.AddNew
        RSTTRXFILE!RCPT_NO = Val(txtBillNo.Text)
        RSTTRXFILE!TRX_TYPE = "PY"
        RSTTRXFILE!INV_NO = cmbinv.Text
        RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMOUNT = Val(txtrcptamt.Text)
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!RCPT_ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From CRDTPYMT WHERE INV_NO = " & Val(cmbinv.Text) & " AND TRX_TYPE='CR'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!RCPT_AMOUNT = RSTTRXFILE!RCPT_AMOUNT + Val(txtrcptamt.Text)
        RSTTRXFILE!BAL_AMT = RSTTRXFILE!INV_AMT - RSTTRXFILE!RCPT_AMOUNT
        If RSTTRXFILE!BAL_AMT <= 0 Then RSTTRXFILE!CHECK_FLAG = "Y" Else RSTTRXFILE!CHECK_FLAG = "N"
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'PY'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    cmbinv.Text = ""
    txtrcptamt.Text = ""
    DataList2.Text = ""
    LBLDATE.Caption = Format(Date, "DD/MM/YYYY")
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
    TXTDEALER.Text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    LBLINVDATE.Caption = ""
    LBLBILLAMT.Caption = ""
    lblrcvdamt.Caption = ""
    LBLBALAMT.Caption = ""
    TMPRCPTNO.Caption = ""
    TMPRCPTDATE.Caption = ""
    TMPACTCODE.Caption = ""
    TMPSUPPLIER.Caption = ""
    TMPINVNO.Caption = ""
    TMPPAIDAMT.Caption = ""
    TMPENTRYDATE.Caption = ""
    LBLCOMBOLIST.Caption = ""
    FRMEMAIN.Enabled = False
    txtBillNo.Enabled = True
    CMDEXIT.Enabled = True
    CmdSave.Enabled = False
    CmdCancel.Enabled = False
    txtBillNo.SetFocus
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            CmdSave.Enabled = False
            txtrcptamt.Enabled = True
            txtrcptamt.SetFocus
    End Select
End Sub

Private Sub DataList2_Click()
    Call FILLINVOICE
    TXTDEALER.Text = DataList2.Text
    cmbinv.Text = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "Payment..."
                DataList2.SetFocus
                Exit Sub
            End If
            If Val(LBLCOMBOLIST.Caption) = 0 Then
                MsgBox "No Payment pending against this Supplier", vbOKOnly, "Payment..."
                TXTDEALER.Enabled = True
                DataList2.Enabled = True
                TXTDEALER.SetFocus
                Exit Sub
            End If
            DataList2.Enabled = False
            TXTDEALER.Enabled = False
            cmbinv.Enabled = True
            cmbinv.SetFocus
        Case vbKeyEscape
            TXTDEALER.Enabled = True
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    Dim rstBILL As ADODB.Recordset
    On Error GoTo eRRHAND
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'PY'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    ACT_FLAG = True
    INV_FLAG = True
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    
    CLOSEALL = 1
    Me.Width = 10000
    Me.Height = 4380
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
        If INV_FLAG = False Then INV_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub Label3_Click()

End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Integer

    On Error GoTo eRRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select MAX(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'PY'", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF And TRXMAST.BOF) Then
                i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
                If Val(txtBillNo.Text) > i Then
                    MsgBox "The last Receipt No. is " & i, vbCritical, "BILL..."
                    txtBillNo.Enabled = True
                    txtBillNo.SetFocus
                    Exit Sub
                End If
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
              
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select MIN(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'PY'", db, adOpenStatic, adLockReadOnly
            If Not (TRXMAST.EOF And TRXMAST.BOF) Then
                i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
                If Val(txtBillNo.Text) < i Then
                    MsgBox "Starting Receipt No. is " & i, vbCritical, "BILL..."
                    txtBillNo.Enabled = True
                    txtBillNo.SetFocus
                    Exit Sub
                End If
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
        
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * from TRNXRCPT WHERE TRX_TYPE='PY' AND RCPT_NO = " & Val(txtBillNo.Text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                If MsgBox("Are you Sure you want to EDIT Receipt No. " & Val(txtBillNo.Text), vbYesNo, "PAYMENT...") = vbNo Then Exit Sub
                TXTINVDATE.Text = Format(rstTRXMAST!RCPT_DATE, "DD/MM/YYYY")
                txtrcptamt.Text = rstTRXMAST!RCPT_AMOUNT
                TXTDEALER.Text = rstTRXMAST!ACT_NAME
                LBLDATE.Caption = Format(rstTRXMAST!RCPT_ENTRY_DATE, "DD/MM/YYYY")
                LBLTIME.Caption = Time
                
                TMPRCPTNO.Caption = Val(txtBillNo.Text)
                TMPRCPTDATE.Caption = TXTINVDATE.Text
                TMPACTCODE.Caption = rstTRXMAST!ACT_CODE
                TMPSUPPLIER.Caption = TXTDEALER.Text
                TMPINVNO.Caption = rstTRXMAST!INV_NO
                TMPPAIDAMT.Caption = txtrcptamt.Text
                TMPENTRYDATE.Caption = LBLDATE.Caption
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * from CRDTPYMT WHERE INV_NO = " & rstTRXMAST!INV_NO & "AND TRX_TYPE='CR'", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    RSTTRXFILE!CHECK_FLAG = "N"
                    RSTTRXFILE!RCPT_AMOUNT = RSTTRXFILE!RCPT_AMOUNT - Val(TMPPAIDAMT.Caption)
                    RSTTRXFILE!BAL_AMT = RSTTRXFILE!INV_AMT - RSTTRXFILE!RCPT_AMOUNT
                    RSTTRXFILE.Update
                    CmdSave.Enabled = False
                    CmdCancel.Enabled = True
                    CMDEXIT.Enabled = False
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                DataList2.BoundText = rstTRXMAST!ACT_CODE
                DataList2_Click
                cmbinv.Text = rstTRXMAST!INV_NO
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            db.Execute ("DELETE from [TRNXRCPT] WHERE TRX_TYPE='PY' AND RCPT_NO = " & Val(TMPRCPTNO.Caption) & "")
            txtBillNo.Enabled = False
            FRMEMAIN.Enabled = True
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
    
    
    'Call TXTBILLNO_KeyDown(13, 0)
    Exit Sub
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
        
    End If
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
    End Select

End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
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
                TXTDEALER.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                TXTINVDATE.Enabled = False
                TXTDEALER.Enabled = True
                DataList2.Enabled = True
                TXTDEALER.SetFocus
            End If
        Case vbKeyEscape
            TXTINVDATE.Enabled = False
            cmdcancel_Click
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
            If Val(txtrcptamt.Text) > Val(LBLBALAMT.Caption) Then
                If MsgBox("The Entered Amount Exceeds Balance Amount by Rs. " & Val(txtrcptamt.Text) - Val(LBLBALAMT.Caption), vbYesNo, "PAYMENT...") = vbNo Then
                    txtrcptamt.SetFocus
                    Exit Sub
                End If
            End If
            txtrcptamt.Enabled = False
            CmdSave.Enabled = True
            CmdSave.SetFocus
        Case vbKeyEscape
            txtrcptamt.Enabled = False
            cmbinv.Enabled = True
            cmbinv.SetFocus
    End Select
End Sub

Private Sub txtrcptamt_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
    flagchange.Caption = ""
End Sub

Private Function FILLINVOICE()
    On Error GoTo eRRHAND
    
    Screen.MousePointer = vbHourglass
    Set cmbinv.DataSource = Nothing
    If INV_FLAG = True Then
        INV_REC.Open "Select * From CRDTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND CHECK_FLAG='N' AND TRX_TYPE='CR' ORDER BY CR_NO", db, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    Else
        INV_REC.Close
        INV_REC.Open "Select * From CRDTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND CHECK_FLAG='N' AND TRX_TYPE='CR' ORDER BY CR_NO", db, adOpenStatic, adLockReadOnly
        INV_FLAG = False
    End If
    
    Set Me.cmbinv.RowSource = INV_REC
    cmbinv.ListField = "INV_NO"
    cmbinv.BoundColumn = "CR_NO"
    LBLCOMBOLIST.Caption = INV_REC.RecordCount
    Screen.MousePointer = vbNormal
    Exit Function

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

