VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "crystl32.ocx"
Begin VB.Form FRMBILL 
   Caption         =   "BILL"
   ClientHeight    =   3510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMBILL.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CMBMONTH 
      ForeColor       =   &H000000FF&
      Height          =   330
      ItemData        =   "FRMBILL.frx":030A
      Left            =   2310
      List            =   "FRMBILL.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   390
      Width           =   2265
   End
   Begin VB.TextBox TXTBILLAMT 
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
      Height          =   315
      Left            =   2310
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1245
      Width           =   780
   End
   Begin VB.TextBox TXTBILLNO 
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
      Height          =   315
      Left            =   2325
      MaxLength       =   10
      TabIndex        =   1
      Top             =   795
      Width           =   780
   End
   Begin VB.CommandButton CMDDISPLAY 
      Caption         =   "&PRINT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1950
      TabIndex        =   3
      Top             =   1755
      Width           =   1230
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
      Height          =   435
      Left            =   3300
      TabIndex        =   4
      Top             =   1755
      Width           =   1230
   End
   Begin Crystal.CrystalReport rptbill 
      Left            =   15
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label LBLTOTAL 
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL BILL NOS."
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
      Index           =   0
      Left            =   795
      TabIndex        =   11
      Top             =   2565
      Width           =   2085
   End
   Begin VB.Label LBLBILLNOS 
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
      Height          =   390
      Left            =   2835
      TabIndex        =   10
      Top             =   2505
      Width           =   1695
   End
   Begin VB.Label LBLTRXTOTAL 
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
      Height          =   390
      Left            =   2850
      TabIndex        =   9
      Top             =   2970
      Width           =   1695
   End
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
      Left            =   795
      TabIndex        =   8
      Top             =   3015
      Width           =   2085
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "APPROX. BILL AMT"
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
      Left            =   420
      TabIndex        =   7
      Top             =   1245
      Width           =   1875
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "STARTING BILL NO."
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
      Left            =   435
      TabIndex        =   6
      Top             =   810
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT MONTH"
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
      Left            =   465
      TabIndex        =   5
      Top             =   405
      Width           =   1425
   End
End
Attribute VB_Name = "FRMBILL"
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
            TXTBILLNO.SetFocus
    End Select
End Sub

Private Sub CMDDISPLAY_Click()
    Dim rstBILLDUP As ADODB.Recordset
    Dim rstBILLORG As ADODB.Recordset
    Dim rstTMP As ADODB.Recordset
    
    Dim E_TABLE As String
    Dim i As Integer
    Dim k As Integer
    Dim n As Integer
    Dim P As Double
    Dim X As Integer
    Dim TRX_AMOUNT As Double
    
    Dim FirstDate As Date, LastDate As Date, M_DATE As Date
    Dim FirstSun As Date
    Dim LastDay As Integer, suncount As Integer, Avoidsun As Integer

    
    If CMBMONTH.ListIndex = -1 Then
        MsgBox "SELECT THE MONTH", vbOKOnly, "BILL"
        CMBMONTH.SetFocus
        Exit Sub
    End If
    
    If Val(TXTBILLNO.Text) = 0 Then
        MsgBox "ENTER THE STARTING BILL NO.", vbOKOnly, "BILL"
        TXTBILLNO.SetFocus
        Exit Sub
    End If
    
    If Val(TXTBILLAMT.Text) = 0 Then
        MsgBox "ENTER THE APPROX. BILL AMOUNT.", vbOKOnly, "BILL"
        TXTBILLAMT.SetFocus
        Exit Sub
    End If
    
    FirstDate = "01" & "/" & CMBMONTH.ListIndex + 1 & "/" & Year(Date)
    M_DATE = FirstDate
    LastDate = DateSerial(Year(Date), Month(Date) + 1, 0)
   
    LastDay = Day(LastDate)
    FirstSun = FirstDate + 7 - Weekday(FirstDate, vbMonday)
    If LastDay >= Day(FirstSun) + 28 Then suncount = 5 Else suncount = 4
    
    Avoidsun = LastDay - suncount
    
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    E_TABLE = "TRXFILE" & Format(CMBMONTH.ListIndex + 1, "00")
    
    P = 0
    Set rstBILLORG = New ADODB.Recordset
    rstBILLORG.Open "Select MIN(Val(VCH_NO)) From " & Trim(E_TABLE) & " ", db, adOpenStatic, adLockReadOnly
    If Not (rstBILLORG.EOF And rstBILLORG.BOF) Then
        P = IIf(IsNull(rstBILLORG.Fields(0)), 0, rstBILLORG.Fields(0))
    End If
    rstBILLORG.Close
    Set rstBILLORG = Nothing
    
    If P = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "NO BILLS AVAILABLE", vbOKOnly, "BILL"
        Exit Sub
    End If
    
    i = Val(TXTBILLNO.Text)
    k = 0
    X = 0
    N_Count = 0
    TRX_AMOUNT = 0
    
    db2.Execute ("DELETE * FROM SALEBILL")
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "Select VCH_NO From " & Trim(E_TABLE) & " ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTMP.EOF
        If TRX_AMOUNT > Val(TXTBILLAMT.Text) Then Exit Do
        n = 0
        Set rstBILLORG = New ADODB.Recordset
        Set rstBILLDUP = New ADODB.Recordset
        rstBILLORG.Open "Select VCH_NO, VCH_DATE, LINE_NO, ITEM_NAME, ITEM_CODE, QTY, MRP, SALES_PRICE, VCH_DESC, REF_NO, TRX_TOTAL, EXP_DATE From " & Trim(E_TABLE) & " WHERE VCH_NO =  " & Val(P) & " AND TRX_TOTAL < 250 AND LINE_NO < 10 ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly, adCmdText
        rstBILLDUP.Open "Select * From SALEBILL", db2, adOpenStatic, adLockOptimistic, adCmdText
        With rstBILLDUP
            Do Until rstBILLORG.EOF
                If X > 10 Then
                    If rstBILLORG.RecordCount < 4 Then GoTo SKIP
                End If
                If rstBILLORG.RecordCount < 4 Then X = X + 1
                rstBILLDUP.AddNew
                n = n + 1
                !VCH_NO = i
                !LINE_NO = n
                !ITEM_NAME = rstBILLORG!ITEM_NAME
                !ITEM_CODE = rstBILLORG!ITEM_CODE
                !QTY = rstBILLORG!QTY
                !MRP = rstBILLORG!MRP
                !SALES_PRICE = rstBILLORG!SALES_PRICE
                !PATIENT = rstBILLORG!VCH_DESC
                !DOCTOR = ""
                !REF_NO = rstBILLORG!REF_NO
                !TRX_TOTAL = rstBILLORG!TRX_TOTAL
                !EXP_DATE = Format(rstBILLORG!EXP_DATE, "mm/yy")
                TRX_AMOUNT = !TRX_TOTAL + TRX_AMOUNT
                
                rstBILLDUP.Update
SKIP:
                rstBILLORG.MoveNext
            Loop
        End With
        rstBILLORG.Close
        Set rstBILLORG = Nothing
        rstTMP.MoveNext
        i = i + 1
        P = P + 1
    Loop
    rstTMP.Close
    Set rstTMP = Nothing
    
    P = 0
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "SELECT SUM(TRX_TOTAL) FROM SALEBILL", db2, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rstTMP.EOF And rstTMP.BOF) Then
        P = IIf(IsNull(rstTMP.Fields(0)), 0, rstTMP.Fields(0))
    End If
    rstTMP.Close
    Set rstTMP = Nothing
    
    If P = 0 Then
        Screen.MousePointer = vbDefault
        MsgBox "NO BILLS AVAILABLE", vbOKOnly, "BILL"
        Exit Sub
    End If
    
    TRX_AMOUNT = 0
    k = 0
    i = 0
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "SELECT VCH_DATE, ITEM_CODE, TRX_TOTAL, MFGR FROM SALEBILL ORDER BY VCH_NO, LINE_NO", db2, adOpenStatic, adLockOptimistic, adCmdText

    Do Until rstTMP.EOF
        rstTMP.MoveNext
        If rstTMP.RecordCount <> i + 1 Then
            If TRX_AMOUNT > (P / Avoidsun) - rstTMP!TRX_TOTAL Then
                k = k + 1
                TRX_AMOUNT = 0
            End If
        End If
        rstTMP.MovePrevious
        M_DATE = DateAdd("d", k, FirstDate)
        If Weekday(M_DATE) = 1 And k = LastDay Then
            k = k - 1
        ElseIf Weekday(M_DATE) = 1 And k <> LastDay Then
            k = k + 1
        ElseIf k > LastDay Then
            k = k - 1
        End If
        
        M_DATE = DateAdd("d", k, FirstDate)
        Set rstBILLORG = New ADODB.Recordset
        rstBILLORG.Open "SELECT ITEM_CODE, MANUFACTURER FROM ITEMMAST WHERE ITEM_CODE = '" & rstTMP!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (rstBILLORG.EOF And rstBILLORG.BOF) Then
            rstTMP!MFGR = IIf(IsNull(rstBILLORG!MANUFACTURER), "", rstBILLORG!MANUFACTURER)
        End If
        rstBILLORG.Close
        Set rstBILLORG = Nothing
        
        rstTMP!VCH_DATE = M_DATE
        TRX_AMOUNT = rstTMP!TRX_TOTAL + TRX_AMOUNT
        TRX_GROSS = rstTMP!TRX_TOTAL + TRX_GROSS
        i = i + 1
        rstTMP.Update
        rstTMP.MoveNext
                
    Loop
    rstTMP.Close
    Set rstTMP = Nothing

    LBLBILLNOS.Caption = i & " NOs."
    LBLTRXTOTAL.Caption = "RS. " & P
    rptbill.ReportFileName = App.Path & "\RPTBILL.RPT"
    'rptbill.Formulas(0) = "Company = '" & LBLSHOP.Caption & "'"
    
    Set rstTMP = New ADODB.Recordset
    rstTMP.Open "SELECT * From NONRCVD", db2, adOpenStatic, adLockOptimistic, adCmdText
    rstTMP.Close
    Set rstTMP = Nothing
    
    rptbill.Action = 1
    
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
        CMBMONTH.ListIndex = Month(Date) - 2
    Else
        CMBMONTH.ListIndex = 11
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmCrimedata.Enabled = True
End Sub

Private Sub TXTBILLAMT_Change()
    LBLBILLNOS.Caption = ""
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub TXTBILLAMT_GotFocus()
    TXTBILLAMT.SelStart = 0
    TXTBILLAMT.SelLength = Len(TXTBILLAMT.Text)
End Sub

Private Sub TXTBILLAMT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTBILLAMT.Text) = 0 Then
                TXTBILLAMT.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        End Select
End Sub

Private Sub TXTBILLAMT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_Change()
    LBLBILLNOS.Caption = ""
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub TXTBILLNO_GotFocus()
    TXTBILLNO.SelStart = 0
    TXTBILLNO.SelLength = Len(TXTBILLNO.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTBILLNO.Text) = 0 Then
                TXTBILLNO.SetFocus
                Exit Sub
            End If
            TXTBILLAMT.SetFocus
    End Select
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
