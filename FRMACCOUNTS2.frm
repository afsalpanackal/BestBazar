VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FRMBILLPRINT 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALE REPORT"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5490
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMACCOUNTS.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5490
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   390
      Left            =   1440
      TabIndex        =   5
      Top             =   240
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   688
      _Version        =   393216
      CalendarForeColor=   0
      CalendarTitleForeColor=   16576
      CalendarTrailingForeColor=   255
      Format          =   16449537
      CurrentDate     =   40498
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   765
      Left            =   360
      TabIndex        =   2
      Top             =   735
      Width           =   4560
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL BILL AMOUNT"
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
         Height          =   315
         Index           =   3
         Left            =   75
         TabIndex        =   4
         Top             =   255
         Width           =   1935
      End
      Begin VB.Label LBLTRXTOTAL 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   2100
         TabIndex        =   3
         Top             =   150
         Width           =   2220
      End
   End
   Begin VB.CommandButton CMDDISPLAY 
      Caption         =   "&DISPLAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   2370
      TabIndex        =   0
      Top             =   1710
      Width           =   1200
   End
   Begin VB.CommandButton CMDEXIT 
      Caption         =   "&EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3735
      TabIndex        =   1
      Top             =   1695
      Width           =   1200
   End
   Begin MSComCtl2.DTPicker DTTO 
      Height          =   390
      Left            =   3435
      TabIndex        =   6
      Top             =   240
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   688
      _Version        =   393216
      Format          =   16449537
      CurrentDate     =   40498
   End
   Begin VB.Label LBLTOTAL 
      BackStyle       =   0  'Transparent
      Caption         =   "TO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   270
      Index           =   5
      Left            =   3060
      TabIndex        =   8
      Top             =   315
      Width           =   375
   End
   Begin VB.Label LBLTOTAL 
      BackStyle       =   0  'Transparent
      Caption         =   "FROM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   270
      Index           =   4
      Left            =   720
      TabIndex        =   7
      Top             =   315
      Width           =   555
   End
End
Attribute VB_Name = "FRMBILLPRINT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMBMONTH_Change()
    BLBILLNOS.Caption = ""
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub CMBMONTH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBMONTH.ListIndex = -1 Then
                CMBMONTH.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
    End Select
End Sub

Private Sub CMDDISPLA()
    Dim rstBILLDUP As ADODB.Recordset
    Dim rstBILLORG As ADODB.Recordset
    Dim rstTMP As ADODB.Recordset
    
    Dim i As Double
    Dim F_BILL As Integer
    Dim L_BILL As Integer
    Dim TRX_AMOUNT As Double
    Dim E_TABLE As String
    
    If CMBMONTH.ListIndex = -1 Then
        MsgBox "SELECT THE MONTH", vbOKOnly, "BILL"
        CMBMONTH.SetFocus
        Exit Sub
    End If
    
    LBLBILLNOS.Caption = ""
    TXTFIRSTBILL.Text = ""
    TXTLASTBILL.Text = ""
    LBLTRXTOTAL.Caption = ""
    
    On Error GoTo ErrHand
    TRX_AMOUNT = 0
    Screen.MousePointer = vbHourglass
    E_TABLE = "TRXFILE" & Format(CMBMONTH.ListIndex + 1, "00")
    
    F_BILL = 0
    L_BILL = 0
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select MIN(Val(VCH_NO)) From " & Trim(E_TABLE) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        F_BILL = IIf(IsNull(rstTMP.Fields(0)), 0, rstTMP.Fields(0))
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select MAX(Val(VCH_NO)) From " & Trim(E_TABLE) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        L_BILL = IIf(IsNull(rstTMP.Fields(0)), 0, rstTMP.Fields(0))
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select DISTINCT T_VCH_NO, VCH_NO From TRXMAST WHERE TRX_TYPE='SI' AND T_VCH_NO >=" & F_BILL & " AND T_VCH_NO <=" & L_BILL & " ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
    Do Until rstTMP.EOF
        Set rstBILLDUP = New ADODB.Recordset
        rstBILLDUP.Open "SELECT * From " & Trim(E_TABLE) & " WHERE TRX_TYPE='SI' AND VCH_NO = " & rstTMP!T_VCH_NO & "", db, adOpenStatic, adLockReadOnly
        Do Until rstBILLDUP.EOF
            TRX_AMOUNT = TRX_AMOUNT + rstBILLDUP!TRX_TOTAL
            rstBILLDUP.MoveNext
        Loop
        rstBILLDUP.Close
        Set rstBILLDUP = Nothing
        rstTMP.MoveNext
    Loop
    If rstTMP.RecordCount > 0 Then
        LBLBILLNOS.Caption = rstTMP.RecordCount
        rstTMP.MoveFirst
        TXTFIRSTBILL.Text = rstTMP!VCH_NO
        rstTMP.MoveLast
        TXTLASTBILL.Text = rstTMP!VCH_NO
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")
    
    TRX_AMOUNT = 0
    Set rstBILLORG = New ADODB.Recordset
    rstBILLORG.Open "SELECT * From " & Trim(E_TABLE) & " WHERE TRX_TYPE='SI' AND VCH_NO >=" & F_BILL & " AND VCH_NO <=" & L_BILL & " ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until rstBILLORG.EOF
        TRX_AMOUNT = TRX_AMOUNT + rstBILLORG!TRX_TOTAL
        rstBILLORG.MoveNext
    Loop
    If rstBILLORG.RecordCount > 0 Then
        LBLORGBILLNOS.Caption = rstBILLORG.RecordCount
        rstBILLORG.MoveFirst
        TXTORGFBILL.Text = rstBILLORG!VCH_NO
        rstBILLORG.MoveLast
        TXTORGLBILL.Text = rstBILLORG!VCH_NO
    End If
    
    rstBILLORG.Close
    Set rstBILLORG = Nothing
    LBLORGAMT.Caption = Format(TRX_AMOUNT, "0.00")
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_Click()
    Dim rstBILLTOTAL As ADODB.Recordset
    
    Dim TRX_AMOUNT As Double
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    LBLTRXTOTAL.Caption = ""
    
    On Error GoTo ErrHand
    TRX_AMOUNT = 0
    Screen.MousePointer = vbHourglass
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")
    
    TRX_AMOUNT = 0
    Set rstBILLTOTAL = New ADODB.Recordset
    rstBILLTOTAL.Open "SELECT * From ATRXFILE WHERE [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # AND CD_FLAG = '2' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    'rstBILLTOTAL.Open "SELECT * From ATRXFILE WHERE [VCH_DATE] <=# " & DTTO.Value & " # AND CD_FLAG = '2'", db, adOpenStatic, adLockReadOnly
    'rstBILLTOTAL.Open "SELECT * From ATRXFILE WHERE [VCH_DATE] =# " & FROMDATE & " # AND CD_FLAG = '2'", db, adOpenStatic, adLockReadOnly, adCmdText

    Do Until rstBILLTOTAL.EOF
        TRX_AMOUNT = TRX_AMOUNT + rstBILLTOTAL!VCH_AMOUNT
        rstBILLTOTAL.MoveNext
    Loop
    rstBILLTOTAL.Close
    Set rstBILLTOTAL = Nothing
    
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    Else
        'CMBMONTH.ListIndex = 11
    End If
    Me.Width = 5580
    Me.Height = 3120
    Me.Left = 0
    Me.Top = 0
    DTFROM.Value = Date
    DTTO.Value = Date
    TXTPASSWORD = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub TXTPASSWORD_GotFocus()
    TXTPASSWORD.Text = ""
    TXTPASSWORD.PasswordChar = " "
    TXTPASSWORD.SelStart = 0
    TXTPASSWORD.SelLength = Len(TXTPASSWORD.Text)
End Sub

Private Sub TXTPASSWORD_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMBMONTH.SetFocus
    End Select
End Sub

Private Sub TXTPASSWORD_LostFocus()
    If UCase(TXTPASSWORD.Text) = "SARAKALAM" Then
        TXTPASSWORD = "YEAR " & Year(Date)
        TXTPASSWORD.PasswordChar = ""
        CMBMONTH.SetFocus
        Me.Height = 6000
    Else
        TXTPASSWORD = "YEAR " & Year(Date)
        TXTPASSWORD.PasswordChar = ""
        CMBMONTH.SetFocus
        Me.Height = 3700
    End If
End Sub
