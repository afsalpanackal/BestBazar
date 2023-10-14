VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMRecpt 
   BackColor       =   &H00D2EDBA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt Entry"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11280
   Icon            =   "FrmRecpt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11280
   Begin VB.Frame Frame3 
      BackColor       =   &H00ECEBCE&
      Height          =   780
      Left            =   4035
      TabIndex        =   21
      Top             =   4065
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
         TabIndex        =   9
         Top             =   210
         Width           =   1365
      End
      Begin VB.CommandButton cmdcancel 
         Caption         =   "E&xit"
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
         TabIndex        =   10
         Top             =   210
         Width           =   1230
      End
   End
   Begin VB.Frame frmemain 
      BackColor       =   &H00ECEBCE&
      Height          =   6810
      Left            =   0
      TabIndex        =   14
      Top             =   -15
      Width           =   11295
      Begin VB.Frame Frame5 
         BackColor       =   &H00D3D18F&
         Caption         =   "Actual Address"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   2220
         Left            =   7050
         TabIndex        =   31
         Top             =   180
         Width           =   4140
         Begin VB.TextBox TxtPhone 
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
            Left            =   735
            MaxLength       =   35
            TabIndex        =   35
            Top             =   1770
            Width           =   2925
         End
         Begin VB.TextBox TXTTIN 
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
            Left            =   735
            MaxLength       =   35
            TabIndex        =   32
            Top             =   1365
            Width           =   3345
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone"
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
            Index           =   35
            Left            =   75
            TabIndex        =   36
            Top             =   1785
            Width           =   660
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Tin No."
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
            Index           =   41
            Left            =   75
            TabIndex        =   34
            Top             =   1395
            Width           =   660
         End
         Begin VB.Label lbladdress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            ForeColor       =   &H80000008&
            Height          =   1110
            Left            =   45
            TabIndex        =   33
            Top             =   210
            Width           =   4035
         End
      End
      Begin VB.TextBox TXTDEALER 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   390
         Left            =   2565
         TabIndex        =   0
         Top             =   180
         Width           =   4455
      End
      Begin VB.TextBox TxtCode 
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
         Height          =   390
         Left            =   1080
         TabIndex        =   2
         Top             =   180
         Width           =   1470
      End
      Begin VB.Frame FrmBank 
         BackColor       =   &H00ECEBCE&
         Height          =   1980
         Left            =   60
         TabIndex        =   25
         Top             =   4785
         Visible         =   0   'False
         Width           =   6930
         Begin VB.CheckBox ChkStatus 
            BackColor       =   &H00D2EDBA&
            Caption         =   "Passed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   5325
            TabIndex        =   41
            Top             =   525
            Width           =   1515
         End
         Begin VB.Frame Frame2 
            Caption         =   "Payment Mode"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   4620
            TabIndex        =   37
            Top             =   810
            Width           =   2250
            Begin VB.OptionButton OptNEFT 
               Caption         =   "NEFT / RTGS etc"
               Height          =   195
               Left            =   75
               TabIndex        =   40
               Top             =   750
               Width           =   1770
            End
            Begin VB.OptionButton OptUPI 
               Caption         =   "UPI"
               Height          =   195
               Left            =   75
               TabIndex        =   39
               Top             =   495
               Width           =   1485
            End
            Begin VB.OptionButton optChq 
               Caption         =   "Cheque / Draft"
               Height          =   195
               Left            =   75
               TabIndex        =   38
               Top             =   270
               Value           =   -1  'True
               Width           =   1485
            End
         End
         Begin MSComCtl2.DTPicker DtChqDate 
            Height          =   360
            Left            =   5325
            TabIndex        =   13
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
            Format          =   61800449
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
            Left            =   1080
            MaxLength       =   20
            TabIndex        =   11
            Top             =   210
            Width           =   3510
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1215
            Left            =   1080
            TabIndex        =   12
            Top             =   675
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   26
            Top             =   210
            Width           =   540
         End
      End
      Begin VB.OptionButton OptCash 
         BackColor       =   &H00ECEBCE&
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
         TabIndex        =   7
         Top             =   4410
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton OptBank 
         BackColor       =   &H00ECEBCE&
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
         TabIndex        =   8
         Top             =   4410
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
         TabIndex        =   6
         Top             =   3705
         Width           =   3465
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
         TabIndex        =   3
         Top             =   3270
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
         TabIndex        =   5
         Top             =   3705
         Width           =   1770
      End
      Begin MSMask.MaskEdBox TXTINVDATE 
         Height          =   345
         Left            =   5340
         TabIndex        =   4
         Top             =   3270
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
      Begin MSDataListLib.DataList DataList2 
         Height          =   2490
         Left            =   1080
         TabIndex        =   1
         Top             =   600
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   4392
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
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
         TabIndex        =   24
         Top             =   210
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
         TabIndex        =   23
         Top             =   3750
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
         TabIndex        =   19
         Top             =   3300
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
         TabIndex        =   18
         Top             =   3300
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
         TabIndex        =   17
         Top             =   3750
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
         TabIndex        =   16
         Top             =   3300
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
         TabIndex        =   15
         Top             =   3270
         Width           =   1215
      End
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   495
   End
   Begin VB.Label lblactcode 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "lblactcode"
      Height          =   315
      Left            =   1065
      TabIndex        =   22
      Top             =   3210
      Width           =   1620
   End
   Begin VB.Label lbltmprcptamt 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "rcpt amount"
      Height          =   315
      Left            =   3150
      TabIndex        =   20
      Top             =   3285
      Width           =   1620
   End
End
Attribute VB_Name = "FRMRecpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_AGNT As New ADODB.Recordset
Dim AGNT_FLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RstBill As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "Select the Custmer from the List", vbOKOnly, "Receipt..."
        TXTDEALER.SetFocus
        Exit Sub
    End If
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Proper Date", vbOKOnly, "Receipt..."
        TXTINVDATE.SetFocus
        Exit Sub
    ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
        MsgBox "Enter Proper Invoice Date", vbOKOnly, "Receipt..."
        TXTINVDATE.SetFocus
        Exit Sub
    Else
        TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
    End If
    
    If Val(txtrcptamt.text) = 0 Then
        MsgBox "Enter Payment Amount", vbOKOnly, "Receipt..."
        txtrcptamt.SetFocus
        Exit Sub
    End If
    
    If OptBank.Value = True And CMBDISTI.BoundText = "" Then
        MsgBox "Please Select the Name of Bank", vbOKOnly, "Receipt..."
        CMBDISTI.SetFocus
        Exit Sub
    End If
    
    If OptBank.Value = True And DateValue(DtChqDate.Value) > DateValue(Date) And ChkStatus.Value = 1 Then
        MsgBox "Please check the status of the Cheque", vbOKOnly, "Receipt..."
        ChkStatus.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    db.BeginTrans
    If OptBank.Value = True Then
        Dim TRX_NO As Double
        TRX_NO = 1
        
        Set RstBill = New ADODB.Recordset
        RstBill.Open "Select MAX(TRX_NO) From BANK_TRX WHERE TRX_TYPE = 'CR' AND BILL_TRX_TYPE = 'RT' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' ", db, adOpenForwardOnly
        If Not (RstBill.EOF And RstBill.BOF) Then
            TRX_NO = IIf(IsNull(RstBill.Fields(0)), 1, RstBill.Fields(0) + 1)
        End If
        RstBill.Close
        Set RstBill = Nothing

    
'        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
'                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_TYPE), "", rstTRANX!B_TRX_TYPE)
'                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
'                GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!B_BILL_TRX_TYPE), "", rstTRANX!B_BILL_TRX_TYPE)
'                GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!B_TRX_YEAR), "", rstTRANX!B_TRX_YEAR)
'                GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
                
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From BANK_TRX", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_NO = TRX_NO
        RSTTRXFILE!TRX_TYPE = "CR"
        RSTTRXFILE!BILL_TRX_TYPE = "RT"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
        RSTTRXFILE!BANK_NAME = CMBDISTI.text
        'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
        RSTTRXFILE!TRX_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!TRX_AMOUNT = Val(txtrcptamt.text)
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.text
        'RSTTRXFILE!INV_DATE = LBLDATE.Caption
        RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
        RSTTRXFILE!ENTRY_DATE = Format(LBLDATE.Caption, "DD/MM/YYYY")
        RSTTRXFILE!CHQ_DATE = Format(DtChqDate.Value, "DD/MM/YYYY")
        RSTTRXFILE!BANK_FLAG = "Y"
        If ChkStatus.Value = 0 Then
            RSTTRXFILE!check_flag = "N"
        Else
            RSTTRXFILE!check_flag = "Y"
        End If
        RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.text)
        If optChq.Value = True Then
            RSTTRXFILE!BANK_MODE = "C"
        ElseIf OptUPI.Value = True Then
            RSTTRXFILE!BANK_MODE = "U"
        ElseIf OptNEFT.Value = True Then
            RSTTRXFILE!BANK_MODE = "N"
        Else
            RSTTRXFILE!BANK_MODE = "C"
        End If
        
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
    End If
            
    Dim BillNO As Long
    BillNO = 1
    Set RstBill = New ADODB.Recordset
    RstBill.Open "Select MAX(REC_NO) From DBTPYMT WHERE TRX_TYPE = 'RT' AND '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
    If Not (RstBill.EOF And RstBill.BOF) Then
        BillNO = IIf(IsNull(RstBill.Fields(0)), 1, RstBill.Fields(0) + 1)
    End If
    RstBill.Close
    Set RstBill = Nothing
        
    Set RstBill = New ADODB.Recordset
    RstBill.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
    If Not (RstBill.EOF And RstBill.BOF) Then
        txtBillNo.text = IIf(IsNull(RstBill.Fields(0)), 1, RstBill.Fields(0) + 1)
    End If
    RstBill.Close
    Set RstBill = Nothing
    
    If optCash.Value = True Then
        Dim RECNO, INVNO As Long
        Dim TRXTYPE, INVTRXTYPE, INVTYPE As String
        
        i = 0
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0) + 1)
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        'db.Execute "Delete FROM CASHATRXFILE WHERE TRX_TYPE = 'CR' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'RT'  AND INV_TRX_TYPE = 'RT'"
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST.AddNew
        RSTITEMMAST!REC_NO = i + 1
        
        RSTITEMMAST!INV_TYPE = "RT"
        RSTITEMMAST!INV_TRX_TYPE = "RT"
        RSTITEMMAST!INV_NO = Val(txtBillNo.text)
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!TRX_TYPE = "CR"
        RSTITEMMAST!check_flag = "S"
        RSTITEMMAST!INV_NO = Val(txtBillNo.text)
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.text
        RSTITEMMAST!AMOUNT = Val(txtrcptamt.text)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!BILL_TRX_TYPE = "SI"
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
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "RT"
    RSTTRXFILE!INV_TRX_TYPE = "RT"
    RSTTRXFILE!CR_NO = Val(txtBillNo.text)
    RSTTRXFILE!REC_NO = BillNO
    'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
    RSTTRXFILE!RCPT_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!RCPT_AMT = Val(txtrcptamt.text)
    RSTTRXFILE!ACT_CODE = DataList2.BoundText
    RSTTRXFILE!ACT_NAME = DataList2.text
    RSTTRXFILE!INV_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!REF_NO = Trim(TXTREFNO.text)
    RSTTRXFILE!INV_AMT = Null
    'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
    RSTTRXFILE!INV_NO = 0
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    If OptBank.Value = True Then
        RSTTRXFILE!BANK_FLAG = "Y"
        RSTTRXFILE!B_TRX_TYPE = "CR"
        RSTTRXFILE!B_TRX_NO = TRX_NO
        RSTTRXFILE!B_BILL_TRX_TYPE = "RT"
        RSTTRXFILE!B_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!BANK_CODE = CMBDISTI.BoundText
        RSTTRXFILE!BANK_NAME = CMBDISTI.text
        RSTTRXFILE!CHQ_NO = Trim(TxtChqNo.text)
        RSTTRXFILE!CHQ_DATE = DtChqDate.Value
        RSTTRXFILE!C_TRX_TYPE = Null
        'RSTTRXFILE!C_REC_NO = Null
        RSTTRXFILE!C_INV_TRX_TYPE = Null
        RSTTRXFILE!C_INV_TYPE = Null
        'RSTTRXFILE!C_INV_NO = Null
    Else
        RSTTRXFILE!BANK_FLAG = "N"
        RSTTRXFILE!B_TRX_TYPE = Null
        'RSTTRXFILE!B_TRX_NO = Null
        RSTTRXFILE!B_BILL_TRX_TYPE = Null
        RSTTRXFILE!B_TRX_YEAR = Null
        RSTTRXFILE!BANK_CODE = Null
        RSTTRXFILE!BANK_NAME = ""
        RSTTRXFILE!C_TRX_TYPE = TRXTYPE
        RSTTRXFILE!C_REC_NO = RECNO
        RSTTRXFILE!C_INV_TRX_TYPE = INVTRXTYPE
        RSTTRXFILE!C_INV_TYPE = INVTYPE
        RSTTRXFILE!C_INV_NO = INVNO
    End If
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
    
    txtrcptamt.text = ""
    TXTDEALER.text = ""
    TxtCode.text = ""
    lbldealer.Caption = ""
    flagchange.Caption = ""
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    DtChqDate.Value = Date
    MsgBox "Saved Successfully....", vbOKOnly, "RECEIPT ENTRY"
    
    Set RstBill = New ADODB.Recordset
    RstBill.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
    If Not (RstBill.EOF And RstBill.BOF) Then
        txtBillNo.text = IIf(IsNull(RstBill.Fields(0)), 1, RstBill.Fields(0) + 1)
    End If
    RstBill.Close
    Set RstBill = Nothing
    TXTDEALER.SetFocus
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
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

Private Sub Form_Activate()
      TXTDEALER.SetFocus
End Sub

Private Sub Form_Load()
    Dim RstBill As ADODB.Recordset
    On Error GoTo ErrHand
    
    AGNT_FLAG = True
    ACT_FLAG = True
    
    Set RstBill = New ADODB.Recordset
    RstBill.Open "Select MAX(CR_NO) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
    If Not (RstBill.EOF And RstBill.BOF) Then
        txtBillNo.text = IIf(IsNull(RstBill.Fields(0)), 1, RstBill.Fields(0) + 1)
    End If
    RstBill.Close
    Set RstBill = Nothing
    
    Call fillcombo
     
    LBLDATE.Caption = Date
    TXTINVDATE.text = Format(Date, "dd/mm/yyyy")
    DtChqDate.Value = Date
    'Width = 8900
    'Height = 4485
    'Left = 1000
    'Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If AGNT_FLAG = False Then ACT_AGNT.Close
    If ACT_FLAG = False Then ACT_REC.Close
End Sub

Private Sub OptBank_Click()
    FrmBank.Visible = True
End Sub

Private Sub optCash_Click()
    FrmBank.Visible = False
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                txtrcptamt.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            ElseIf Len(Trim(TXTINVDATE.text)) < 10 Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                txtrcptamt.SetFocus
            End If
    End Select
End Sub

Private Sub txtrcptamt_GotFocus()
    txtrcptamt.SelStart = 0
    txtrcptamt.SelLength = Len(txtrcptamt.text)
End Sub

Private Sub txtrcptamt_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(txtrcptamt.text) = 0 Then
                MsgBox "Enter Payment Amount", vbOKOnly, "RECEIPT..."
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREFNO_GotFocus()
    TXTREFNO.SelStart = 0
    TXTREFNO.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTREFNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdSAVE.SetFocus
    End Select
End Sub

Private Sub TXTREFNO_KeyPress(KeyAscii As Integer)

    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/"), vbKey0 To vbKey9
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select

End Sub

Private Function fillcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If AGNT_FLAG = True Then
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    Else
        ACT_AGNT.Close
        ACT_AGNT.Open "select BANK_CODE, BANK_NAME from BANKCODE ORDER BY BANK_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        AGNT_FLAG = False
    End If
    Set Me.CMBDISTI.RowSource = ACT_AGNT
    CMBDISTI.ListField = "BANK_NAME"
    CMBDISTI.BoundColumn = "BANK_CODE"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TXTCODE_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
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

Private Sub DataList2_Click()
    Dim rstCustomer As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    On Error GoTo ErrHand
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = DataList2.text & Chr(13) & Trim(rstCustomer!Address)
        TXTTIN.text = IIf(IsNull(rstCustomer!KGST), "", Trim(rstCustomer!KGST))
        TxtPhone.text = IIf(IsNull(rstCustomer!TELNO), "", Trim(rstCustomer!TELNO))
    Else
        TxtPhone.text = ""
        TXTTIN.text = ""
        lbladdress.Caption = ""
    End If
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
    TxtCode.text = DataList2.BoundText
    Exit Sub
    
ErrHand:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    TxtCode.text = DataList2.BoundText
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

