VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMDRCR 
   BackColor       =   &H00C000C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Debit Note / Credit Note Entry"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "FrmDrCr.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   7050
   Begin VB.Frame Frame3 
      BackColor       =   &H00D2EDBA&
      Height          =   780
      Left            =   4035
      TabIndex        =   10
      Top             =   2475
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
      BackColor       =   &H00D2EDBA&
      Height          =   5250
      Left            =   0
      TabIndex        =   3
      Top             =   -15
      Width           =   7050
      Begin VB.Frame Frame1 
         BackColor       =   &H00D2EDBA&
         Height          =   885
         Left            =   75
         TabIndex        =   29
         Top             =   105
         Width           =   6930
         Begin VB.OptionButton OptDr 
            BackColor       =   &H00D2EDBA&
            Caption         =   "Debit Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000707F1&
            Height          =   345
            Left            =   3090
            TabIndex        =   31
            Top             =   150
            Width           =   2340
         End
         Begin VB.OptionButton OptCr 
            BackColor       =   &H00D2EDBA&
            Caption         =   "Credit Note"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000707F1&
            Height          =   345
            Left            =   315
            TabIndex        =   30
            Top             =   150
            Value           =   -1  'True
            Width           =   2340
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "(Discount, Sales Return etc.)"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   300
            Index           =   9
            Left            =   75
            TabIndex        =   32
            Top             =   510
            Width           =   3000
         End
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00D2EDBA&
         Caption         =   "Entry"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   75
         TabIndex        =   28
         Top             =   2775
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.Frame FrmBank 
         BackColor       =   &H00D2EDBA&
         Height          =   1980
         Left            =   60
         TabIndex        =   20
         Top             =   3195
         Visible         =   0   'False
         Width           =   6930
         Begin MSComCtl2.DTPicker DtChqDate 
            Height          =   360
            Left            =   5325
            TabIndex        =   26
            Top             =   165
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   635
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   82837505
            CurrentDate     =   41452
         End
         Begin VB.CheckBox ChkStatus 
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   4860
            TabIndex        =   25
            Top             =   660
            Width           =   1890
         End
         Begin VB.TextBox TxtChqNo 
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
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   22
            Top             =   210
            Width           =   3510
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1215
            Left            =   1080
            TabIndex        =   27
            Top             =   660
            Width           =   3510
            _ExtentX        =   6191
            _ExtentY        =   2143
            _Version        =   393216
            Appearance      =   0
            Style           =   1
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Bank"
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
            Index           =   7
            Left            =   105
            TabIndex        =   24
            Top             =   705
            Width           =   645
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Cheque / Draft No."
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
            Height          =   435
            Index           =   6
            Left            =   90
            TabIndex        =   23
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Index           =   5
            Left            =   4770
            TabIndex        =   21
            Top             =   210
            Width           =   540
         End
      End
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00D2EDBA&
         Caption         =   "By Cash"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   1275
         TabIndex        =   19
         Top             =   2775
         Width           =   1410
      End
      Begin VB.OptionButton OptBank 
         BackColor       =   &H00D2EDBA&
         Caption         =   "To Bank"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   270
         Left            =   2760
         TabIndex        =   18
         Top             =   2775
         Width           =   1230
      End
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
         Left            =   3525
         MaxLength       =   20
         TabIndex        =   13
         Top             =   2100
         Width           =   3405
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
         Top             =   1560
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
         Left            =   975
         MaxLength       =   8
         TabIndex        =   2
         Top             =   2130
         Width           =   1770
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5340
         TabIndex        =   15
         Top             =   1560
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
         Top             =   1005
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
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
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
         Left            =   2910
         TabIndex        =   14
         Top             =   2130
         Width           =   645
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Trnx"
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
         Top             =   1560
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
         Top             =   1560
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
         Top             =   2130
         Width           =   1035
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
         Top             =   1560
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
         Top             =   1560
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
Attribute VB_Name = "FRMDRCR"
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
    Dim i As Integer
    
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
    
    If OptBank.value = True And CMBDISTI.BoundText = "" Then
        MsgBox "Please Select the Name of Bank", vbOKOnly, "Receipt..."
        CMBDISTI.SetFocus
        Exit Sub
    End If
    
    If OptBank.value = True And DateValue(DtChqDate.value) > DateValue(Date) And ChkStatus.value = 1 Then
        MsgBox "Please check the status of the Cheque", vbOKOnly, "Receipt..."
        ChkStatus.SetFocus
        Exit Sub
    End If
    
    On Error GoTo errHand
    
    If OptBank.value = True Then
        Dim TRX_NO As Double
        TRX_NO = 1
        
        Set rstBILL = New ADODB.Recordset
        If OptCr.value = True Then
            rstBILL.Open "Select MAX(Val(TRX_NO)) From BANK_TRX WHERE TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'CN' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
        Else
            rstBILL.Open "Select MAX(Val(TRX_NO)) From BANK_TRX WHERE TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'DN' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
        End If
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            TRX_NO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing

        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From BANK_TRX", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_NO = TRX_NO
        If OptCr.value = True Then
            RSTTRXFILE!TRX_TYPE = "CR"
            RSTTRXFILE!BILL_TRX_TYPE = "CN"
        Else
            RSTTRXFILE!TRX_TYPE = "DR"
            RSTTRXFILE!BILL_TRX_TYPE = "DN"
        End If
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
        RSTTRXFILE!BANK_NAME = CMBDISTI.Text
        'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
        RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.Text)
        RSTTRXFILE!ACT_CODE = lblactcode.Caption
        RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
        'RSTTRXFILE!INV_DATE = LBLDATE.Caption
        RSTTRXFILE!REF_NO = Trim(TXTREFNO.Text)
        RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
        RSTTRXFILE!CHQ_DATE = Format(DtChqDate.value, "DD/MM/YYYY")
        RSTTRXFILE!BANK_FLAG = "Y"
        If ChkStatus.value = 0 Then
            RSTTRXFILE!CHECK_FLAG = "N"
        Else
            RSTTRXFILE!CHECK_FLAG = "Y"
        End If
        RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.Text)
        
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
    
    Set rstBILL = New ADODB.Recordset
    If OptCr.value = True Then
        rstBILL.Open "Select MAX(Val(CR_NO)) From DBTPYMT WHERE TRX_TYPE = 'CB'", db, adOpenForwardOnly
    Else
        rstBILL.Open "Select MAX(Val(CR_NO)) From DBTPYMT WHERE TRX_TYPE = 'DB'", db, adOpenForwardOnly
    End If
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    If OptCash.value = True Then
        Dim RECNO, INVNO As Long
        Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
        
        i = 0
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(Val(REC_NO)) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0) + 1)
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
       
        If OptCr.value = True Then
            db.Execute "delete * FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'CN'  AND INV_TRX_TYPE = 'CN'"
        Else
            db.Execute "delete * FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'DN'  AND INV_TRX_TYPE = 'DN'"
        End If

        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST.AddNew
        RSTITEMMAST!rec_no = i + 1
        If OptCr.value = True Then
            RSTITEMMAST!INV_TYPE = "CN"
            RSTITEMMAST!INV_TRX_TYPE = "CN"
            RSTITEMMAST!CHECK_FLAG = "S"
        Else
            RSTITEMMAST!INV_TYPE = "DN"
            RSTITEMMAST!INV_TRX_TYPE = "DN"
            RSTITEMMAST!CHECK_FLAG = "P"
        End If
        RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
        RSTITEMMAST!TRX_TYPE = "CR"
        RSTITEMMAST!ACT_CODE = lblactcode.Caption
        RSTITEMMAST!ACT_NAME = LBLSUPPLIER.Caption
        RSTITEMMAST!AMOUNT = Val(txtrcptamt.Text)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!BILL_TRX_TYPE = "SI"
        RSTITEMMAST!CASH_MODE = "C"
        RSTITEMMAST!CHQ_NO = ""
        RSTITEMMAST!CHQ_DATE = Null
        RSTITEMMAST!BANK = ""
        RSTITEMMAST!CHQ_STATUS = ""
        
        RECNO = RSTITEMMAST!rec_no
        INVNO = RSTITEMMAST!INV_NO
        TRXTYPE = RSTITEMMAST!TRX_TYPE
        INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
        INVTYPE = RSTITEMMAST!INV_TYPE
        
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    If OptCr.value = True Then
        RSTTRXFILE!TRX_TYPE = "CB"
        RSTTRXFILE!INV_TRX_TYPE = "CN"
    Else
        RSTTRXFILE!TRX_TYPE = "DB"
        RSTTRXFILE!INV_TRX_TYPE = "DN"
    End If
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
    If OptBank.value = True Then
        If OptCr.value = True Then
            RSTTRXFILE!B_TRX_TYPE = "CB"
            RSTTRXFILE!B_BILL_TRX_TYPE = "CN"
        Else
            RSTTRXFILE!B_TRX_TYPE = "DB"
            RSTTRXFILE!B_BILL_TRX_TYPE = "DN"
        End If
        RSTTRXFILE!BANK_FLAG = "Y"
        RSTTRXFILE!B_TRX_NO = TRX_NO
        RSTTRXFILE!B_TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
        
        RSTTRXFILE!C_TRX_TYPE = Null
        RSTTRXFILE!C_REC_NO = Null
        RSTTRXFILE!C_INV_TRX_TYPE = Null
        RSTTRXFILE!C_INV_TYPE = Null
        RSTTRXFILE!C_INV_NO = Null
    Else
        RSTTRXFILE!BANK_FLAG = "N"
        RSTTRXFILE!B_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_NO = Null
        RSTTRXFILE!B_BILL_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_YEAR = Null
        RSTTRXFILE!BANK_CODE = Null
        
        RSTTRXFILE!C_TRX_TYPE = TRXTYPE
        RSTTRXFILE!C_REC_NO = RECNO
        RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
        RSTTRXFILE!C_INV_TYPE = INVTYPE
        RSTTRXFILE!C_INV_NO = INVNO
    End If
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    MsgBox "Saved Successfully....", vbOKOnly, "PAYMENT"
    Unload Me
    Exit Sub
errHand:
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
    On Error GoTo errHand
    
    AGNT_FLAG = True
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(Val(CR_NO)) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Call FILLCOMBO
     
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    DtChqDate.value = Date
    'Width = 8900
    'Height = 4485
    Left = 800
    Top = 1000
    
    Exit Sub
errHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
    FRMRcptReg.Enabled = True
    FRMRcptReg.GRDTranx.SetFocus
End Sub

Private Sub OptBank_Click()
    FrmBank.Visible = True
End Sub

Private Sub OptCash_Click()
    FrmBank.Visible = False
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
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
            cmdSAVE.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Function FILLCOMBO()
    On Error GoTo errHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from [BANKCODE] ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from [BANKCODE] ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "BANK_NAME"
    CMBDISTI.BoundColumn = "BANK_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

errHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

