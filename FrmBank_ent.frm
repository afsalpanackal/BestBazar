VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMbankentry 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bank Entries"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   ControlBox      =   0   'False
   Icon            =   "FrmBank_ent.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7050
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1530
      Left            =   0
      TabIndex        =   22
      Top             =   -60
      Width           =   7035
      Begin VB.OptionButton OptCharge 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Bank Charges"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   4215
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   31
         Top             =   945
         Width           =   2175
      End
      Begin VB.OptionButton OptWithdraw 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Withdrawal"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   4215
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   30
         Top             =   390
         Width           =   1770
      End
      Begin VB.OptionButton OptDeposit 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Deposit"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   180
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   28
         Top             =   375
         Width           =   1740
      End
      Begin VB.OptionButton OptInterst 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Bank Intersts"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   180
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   29
         Top             =   930
         Width           =   2115
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Debit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   4035
         TabIndex        =   24
         Top             =   135
         Width           =   2970
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Credit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   60
         TabIndex        =   23
         Top             =   135
         Width           =   2970
      End
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00C0FFC0&
      Height          =   4275
      Left            =   0
      TabIndex        =   3
      Top             =   1395
      Visible         =   0   'False
      Width           =   7050
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Left            =   255
         TabIndex        =   1
         Top             =   2460
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton OptBank 
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Left            =   1935
         TabIndex        =   17
         Top             =   2460
         Width           =   1410
      End
      Begin VB.TextBox TXTREFNO 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   405
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   12
         Top             =   1845
         Width           =   5640
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
         TabIndex        =   10
         Top             =   765
         Width           =   795
      End
      Begin VB.TextBox txtrcptamt 
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
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Left            =   2115
         MaxLength       =   8
         TabIndex        =   0
         Top             =   1350
         Width           =   1785
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5430
         TabIndex        =   14
         Top             =   765
         Width           =   1485
         _ExtentX        =   2619
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
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   3675
         TabIndex        =   25
         Top             =   2220
         Width           =   3225
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
            Height          =   435
            Left            =   1800
            TabIndex        =   27
            Top             =   210
            Width           =   1290
         End
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
            Height          =   435
            Left            =   105
            TabIndex        =   26
            Top             =   210
            Width           =   1410
         End
      End
      Begin VB.Frame FrmBank 
         BackColor       =   &H00C0FFC0&
         Height          =   1335
         Left            =   60
         TabIndex        =   18
         Top             =   2895
         Visible         =   0   'False
         Width           =   6930
         Begin VB.CheckBox ChkStatus 
            BackColor       =   &H00D2EDBA&
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   5340
            TabIndex        =   36
            Top             =   555
            Width           =   1515
         End
         Begin MSComCtl2.DTPicker DtChqDate 
            Height          =   360
            Left            =   5325
            TabIndex        =   21
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
            Format          =   125042689
            CurrentDate     =   41452
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
            Left            =   1350
            MaxLength       =   20
            TabIndex        =   2
            Top             =   210
            Width           =   3345
         End
         Begin VB.Frame Frame5 
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   630
            Left            =   75
            TabIndex        =   32
            Top             =   675
            Width           =   4200
            Begin VB.OptionButton OptNEFT 
               Caption         =   "NEFT / RTGS etc"
               Height          =   195
               Left            =   2310
               TabIndex        =   35
               Top             =   285
               Width           =   1770
            End
            Begin VB.OptionButton OptUPI 
               Caption         =   "UPI"
               Height          =   195
               Left            =   1515
               TabIndex        =   34
               Top             =   285
               Width           =   1485
            End
            Begin VB.OptionButton optChq 
               Caption         =   "Cheque / Draft"
               Height          =   195
               Left            =   75
               TabIndex        =   33
               Top             =   285
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Trnx / Ref No."
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
            TabIndex        =   20
            Top             =   135
            Width           =   1215
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
            TabIndex        =   19
            Top             =   210
            Width           =   540
         End
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
         Left            =   660
         TabIndex        =   16
         Top             =   195
         Width           =   4035
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   4
         Left            =   105
         TabIndex        =   15
         Top             =   225
         Width           =   600
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
         Height          =   360
         Index           =   8
         Left            =   105
         TabIndex        =   13
         Top             =   1935
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Date of Tranx"
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
         Width           =   2085
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
         Caption         =   "Rcvd Amount"
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
         Top             =   1410
         Width           =   1860
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
      TabIndex        =   11
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
Attribute VB_Name = "FRMbankentry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

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
        MsgBox "Enter Proper Date", vbOKOnly, "Transaction..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.Text)) < 10 Then
        MsgBox "Enter Proper Date", vbOKOnly, "Transaction..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    End If
    
    If Val(txtrcptamt.Text) = 0 Then
        MsgBox "Enter the Transaction Amount", vbOKOnly, "Transaction..."
        txtrcptamt.SetFocus
        Exit Sub
    End If
    
    If Trim(TXTREFNO.Text) = "" Then
        MsgBox "Please enter a remarks for the Transaction", vbOKOnly, "Transaction..."
        TXTREFNO.SetFocus
        Exit Sub
    End If
    
    If OptBank.value = True And DateValue(DtChqDate.value) > DateValue(Date) And ChkStatus.value = 1 Then
        MsgBox "Please check the status of the Cheque", vbOKOnly, "Receipt..."
        ChkStatus.SetFocus
        Exit Sub
    End If
    
    On Error GoTo Errhand
    db.BeginTrans
    If OptCash.value = True Then
        If OptDeposit.value = True Or OptWithdraw.value = True Then
            Dim RECNO, INVNO As Long
            Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
'            If OptDeposit.value = True Then
'                db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'DP'  AND INV_TRX_TYPE = 'DP' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "'"
'            ElseIf OptWithdraw.value = True Then
'                db.Execute "delete FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'WD'  AND INV_TRX_TYPE = 'WD' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "'"
'            End If
            
            i = 0
            Set rstMaxRec = New ADODB.Recordset
            rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
            If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
                i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0) + 1)
            End If
            rstMaxRec.Close
            Set rstMaxRec = Nothing

            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'WD'  AND INV_TRX_TYPE = 'WD' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            RSTITEMMAST.AddNew
            RSTITEMMAST!REC_NO = i + 1
            If OptDeposit.value = True Then
                RSTITEMMAST!INV_TYPE = "DP"
                RSTITEMMAST!INV_TRX_TYPE = "DP"
                RSTITEMMAST!CHECK_FLAG = "P"
            ElseIf OptWithdraw.value = True Then
                RSTITEMMAST!INV_TYPE = "WD"
                RSTITEMMAST!INV_TRX_TYPE = "WD"
                RSTITEMMAST!CHECK_FLAG = "S"
            End If
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
            RSTITEMMAST!TRX_TYPE = "CR"
            RSTITEMMAST!ACT_CODE = lblactcode.Caption
            RSTITEMMAST!ACT_NAME = LBLSUPPLIER.Caption
            RSTITEMMAST!AMOUNT = Val(txtrcptamt.Text)
            RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
            RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
            RSTITEMMAST!BILL_TRX_TYPE = ""
            RSTITEMMAST!CASH_MODE = "C"
            RSTITEMMAST!CHQ_NO = ""
            'RSTITEMMAST!CHQ_DATE = Null
            RSTITEMMAST!BANK = ""
            RSTITEMMAST!CHQ_STATUS = ""
            
            RECNO = RSTITEMMAST!REC_NO
            INVNO = RSTITEMMAST!INV_NO
            TRXTYPE = RSTITEMMAST!TRX_TYPE
            INVTRXTYPE = RSTITEMMAST!INV_TRX_TYPE
            INVTYPE = RSTITEMMAST!INV_TYPE
            
            RSTITEMMAST.Update
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        End If
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From BANK_TRX", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_NO = Val(txtBillNo.Text)
    If OptDeposit.value = True Then
        RSTTRXFILE!TRX_TYPE = "CR"
        RSTTRXFILE!BILL_TRX_TYPE = "DP"
    ElseIf OptInterst.value = True Then
        RSTTRXFILE!TRX_TYPE = "CR"
        RSTTRXFILE!BILL_TRX_TYPE = "IN"
    ElseIf OptWithdraw.value = True Then
        RSTTRXFILE!TRX_TYPE = "DR"
        RSTTRXFILE!BILL_TRX_TYPE = "WD"
    Else
        RSTTRXFILE!TRX_TYPE = "DR"
        RSTTRXFILE!BILL_TRX_TYPE = "BC"
    End If
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
    RSTTRXFILE!BANK_CODE = lblactcode.Caption
    RSTTRXFILE!BANK_NAME = LBLSUPPLIER.Caption
    'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
    RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.Text)
    RSTTRXFILE!ACT_CODE = lblactcode.Caption
    RSTTRXFILE!ACT_NAME = LBLSUPPLIER.Caption
    'RSTTRXFILE!INV_DATE = LBLDATE.Caption
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.Text)
    RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
    If OptBank.value = True Then
        RSTTRXFILE!BANK_FLAG = "Y"
        RSTTRXFILE!CHQ_DATE = Format(DtChqDate.value, "DD/MM/YYYY")
        If ChkStatus.value = 0 Then
            RSTTRXFILE!CHECK_FLAG = "N"
        Else
            RSTTRXFILE!CHECK_FLAG = "Y"
        End If
        RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.Text)
        If optChq.value = True Then
            RSTTRXFILE!BANK_MODE = "C"
        ElseIf OptUPI.value = True Then
            RSTTRXFILE!BANK_MODE = "U"
        ElseIf OptNEFT.value = True Then
            RSTTRXFILE!BANK_MODE = "N"
        Else
            RSTTRXFILE!BANK_MODE = "C"
        End If
        RSTTRXFILE!C_TRX_TYPE = Null
        'RSTTRXFILE!C_REC_NO = Null
        RSTTRXFILE!C_INV_TRX_TYPE = Null
        RSTTRXFILE!C_INV_TYPE = Null
        If optChq.value = True Then
            RSTTRXFILE!BANK_MODE = "C"
        ElseIf OptUPI.value = True Then
            RSTTRXFILE!BANK_MODE = "U"
        ElseIf OptNEFT.value = True Then
            RSTTRXFILE!BANK_MODE = "N"
        Else
            RSTTRXFILE!BANK_MODE = "C"
        End If
        ''RSTTRXFILE!C_INV_NO = Null
    Else
        RSTTRXFILE!BANK_FLAG = "N"
        'RSTTRXFILE!CHQ_DATE = Null
        RSTTRXFILE!CHECK_FLAG = ""
        RSTTRXFILE!CHQ_NO = ""
        If OptDeposit.value = True Or OptWithdraw.value = True Then
            RSTTRXFILE!C_TRX_TYPE = TRXTYPE
            RSTTRXFILE!C_REC_NO = RECNO
            RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
            RSTTRXFILE!C_INV_TYPE = INVTYPE
            RSTTRXFILE!C_INV_NO = INVNO
        Else
            RSTTRXFILE!C_TRX_TYPE = Null
            'RSTTRXFILE!C_REC_NO = Null
            RSTTRXFILE!C_INV_TRX_TYPE = Null
            RSTTRXFILE!C_INV_TYPE = Null
            ''RSTTRXFILE!C_INV_NO = Null
        End If
        RSTTRXFILE!BANK_MODE = ""
    End If
    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
    
    
   'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'RT'  AND INV_TRX_TYPE = 'RT'"
    
'    i = 0
'    Set rstMaxRec = New ADODB.Recordset
'    rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
'    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
'        i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
'    End If
'    rstMaxRec.Close
'    Set rstMaxRec = Nothing
'
'    'db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(creditbill.LBLBILLNO.Caption) & " AND INV_TYPE = 'RT'"
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
'    RSTITEMMAST.AddNew
'    RSTITEMMAST!REC_NO = i + 1
'    RSTITEMMAST!INV_TYPE = "RT"
'    RSTITEMMAST!INV_TRX_TYPE = "RT"
'    RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
'    RSTITEMMAST!TRX_TYPE = "CR"
'    RSTITEMMAST!ACT_CODE = lblactcode.Caption
'    RSTITEMMAST!ACT_NAME = LBLSUPPLIER.Caption
'    RSTITEMMAST!AMOUNT = Val(txtrcptamt.Text)
'    RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
'    RSTITEMMAST!CHECK_FLAG = "S"
'    RSTITEMMAST!BILL_TRX_TYPE = "SI"
'    If OptCash.value = True Then
'        RSTITEMMAST!CASH_MODE = "C"
'        RSTITEMMAST!CHQ_NO = ""
'        RSTITEMMAST!CHQ_DATE = Null
'        RSTITEMMAST!BANK = ""
'        RSTITEMMAST!CHQ_STATUS = ""
'    Else
'        RSTITEMMAST!CASH_MODE = "B"
'        RSTITEMMAST!CHQ_NO = Trim(TxtChqNo.Text)
'        RSTITEMMAST!CHQ_DATE = DtChqDate.value
'        RSTITEMMAST!BANK = Trim(TxtChqNo.Text)
'        If ChkStatus.value = 1 Then
'            RSTITEMMAST!CHQ_STATUS = "Y"
'        Else
'            RSTITEMMAST!CHQ_STATUS = "N"
'        End If
'    End If
'    RSTITEMMAST.Update
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
'
    
    
    MsgBox "Saved Successfully....", vbOKOnly, "PAYMENT"
    Unload Me
    Exit Sub
Errhand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtrcptamt.SetFocus
    End Select
End Sub

Private Sub Form_Load()
     
    LBLDATE.Caption = Date
    TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    DtChqDate.value = Date
    'Width = 8900
    'Height = 4485
    Left = 800
    Top = 1000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FRMBankBook.Enabled = True
    FRMBankBook.GRDTranx.SetFocus
End Sub

Private Sub OptBank_Click()
    FrmBank.Visible = True
End Sub

Private Sub OptCash_Click()
    FrmBank.Visible = False
End Sub

Private Sub OptCharge_Click()
    Dim rstBILL As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BANK_CODE='" & lblactcode.Caption & "' AND  TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'BC' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    frmemain.Visible = True
    txtrcptamt.SetFocus
    Label1(2).Caption = "Bank Charges"
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub OptDeposit_Click()
    Dim rstBILL As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BANK_CODE='" & lblactcode.Caption & "' AND TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'DP' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    frmemain.Visible = True
    txtrcptamt.SetFocus
    
    Label1(2).Caption = "Deposit Amt"
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub OptInterst_Click()
    Dim rstBILL As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BANK_CODE='" & lblactcode.Caption & "' AND  TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'IN' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    frmemain.Visible = True
    txtrcptamt.SetFocus
    Label1(2).Caption = "Bank Interests"
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub OptWithdraw_Click()
    Dim rstBILL As ADODB.Recordset
    
    On Error GoTo Errhand
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(TRX_NO) From BANK_TRX WHERE BANK_CODE='" & lblactcode.Caption & "' AND  TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'WD' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' ", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    frmemain.Visible = True
    txtrcptamt.SetFocus
    Label1(2).Caption = "Withdrawal Amt"
    Exit Sub
Errhand:
    MsgBox Err.Description
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
            cmdSAVE.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(Chr(KeyAscii))
        Case Else
            KeyAscii = 0
    End Select

End Sub