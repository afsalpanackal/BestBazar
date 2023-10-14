VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSrvmovmnt 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SERVICES REGISTER"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11070
   ClipControls    =   0   'False
   Icon            =   "frmsrvmvmt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   11070
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9660
      TabIndex        =   24
      Top             =   7830
      Width           =   1380
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Caption         =   "SORT ORDER"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   945
      Left            =   6255
      TabIndex        =   12
      Top             =   525
      Width           =   4785
      Begin VB.OptionButton OPTOUTDATE 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   105
         TabIndex        =   14
         Top             =   225
         Value           =   -1  'True
         Width           =   3285
      End
      Begin VB.OptionButton OPTOUTCUST 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   555
         Width           =   3285
      End
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
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   6195
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
      Height          =   405
      Left            =   8130
      TabIndex        =   2
      Top             =   7830
      Width           =   1380
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1035
      Left            =   45
      TabIndex        =   1
      Top             =   435
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   1826
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
   Begin MSFlexGridLib.MSFlexGrid GRDOUTWARD 
      Height          =   5670
      Left            =   0
      TabIndex        =   7
      Top             =   1770
      Width           =   11040
      _ExtentX        =   19473
      _ExtentY        =   10001
      _Version        =   393216
      Rows            =   1
      Cols            =   10
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   8438015
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   345
      Left            =   7650
      TabIndex        =   18
      Top             =   105
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      _Version        =   393216
      CalendarForeColor=   0
      CalendarTitleForeColor=   16576
      CalendarTrailingForeColor=   255
      Format          =   116391937
      CurrentDate     =   41640
      MinDate         =   40179
   End
   Begin MSComCtl2.DTPicker DTTO 
      Height          =   345
      Left            =   9525
      TabIndex        =   19
      Top             =   120
      Width           =   1530
      _ExtentX        =   2699
      _ExtentY        =   609
      _Version        =   393216
      Format          =   116391937
      CurrentDate     =   41640
      MinDate         =   41640
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OP. QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   4
      Left            =   120
      TabIndex        =   23
      Top             =   8280
      Visible         =   0   'False
      Width           =   1365
   End
   Begin VB.Label LBLOPQTY 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   1545
      TabIndex        =   22
      Top             =   8280
      Visible         =   0   'False
      Width           =   1500
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
      ForeColor       =   &H00FF0000&
      Height          =   225
      Index           =   5
      Left            =   9195
      TabIndex        =   21
      Top             =   150
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "For the Period"
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
      Left            =   6270
      TabIndex        =   20
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   3
      Left            =   6480
      TabIndex        =   17
      Top             =   8325
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOOSE QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   2
      Left            =   3000
      TabIndex        =   16
      Top             =   8325
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label LblLoose 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   4860
      TabIndex        =   15
      Top             =   8265
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label lblmanual 
      Caption         =   "*Adjusted Manually"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   10800
      TabIndex        =   11
      Top             =   8205
      Visible         =   0   'False
      Width           =   2355
   End
   Begin VB.Label LBLHEAD 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Index           =   1
      Left            =   15
      TabIndex        =   10
      Top             =   1515
      Width           =   11025
   End
   Begin VB.Label LBLOUTWARD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   4860
      TabIndex        =   9
      Top             =   7485
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "AMOUNT"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   345
      Index           =   1
      Left            =   6435
      TabIndex        =   8
      Top             =   7500
      Width           =   1080
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "OUTWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   0
      Left            =   3180
      TabIndex        =   6
      Top             =   7530
      Width           =   1785
   End
   Begin VB.Label LBLBALANCE 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   7485
      TabIndex        =   5
      Top             =   7455
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "INWARD QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Index           =   6
      Left            =   120
      TabIndex        =   4
      Top             =   7875
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Label LBLINWARD 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   1560
      TabIndex        =   3
      Top             =   7875
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "FrmSrvmovmnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean 'REP
Dim RSTREP As New ADODB.Recordset

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ErrHand
    ReportNameVar = Rptpath & "RPTREPORT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    'Report.OpenSubreport("Rptinward").RecordSelectionFormula = "( ({TRXFILE.TRX_TYPE} = 'OG' OR {TRXFILE.TRX_TYPE} = 'PI' OR {TRXFILE.TRX_TYPE} = 'OP') AND {TRXFILE.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({RTRXFILE.ITEM_CODE} = '" & DataList2.BoundText & "' AND {RTRXFILE.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {RTRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    'Report.OpenSubreport("RPTINWRD.rpt").RecordSelectionFormula = "({TRANSMAST.VCH_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {TRANSMAST.VCH_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTINWRD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTINWRD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTINWRD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTOUTWARD.rpt").RecordSelectionFormula = "({TRXFILE.ITEM_CODE} = '" & DataList2.BoundText & "' AND {TRXMAST.VCH_DATE} <=# " & Format(DTTO.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    For i = 1 To Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables.COUNT
        Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.OpenSubreport("RPTOUTWARD.rpt").Database.Tables(i).Name & " ")
            Report.OpenSubreport("RPTOUTWARD.rpt").Database.SetDataSource oRs, 3, i
            Set oRs = Nothing
        End If
    Next i
    Report.OpenSubreport("RPTINWRD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTINWRD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTINWRD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    
    Report.OpenSubreport("RPTOUTWARD.rpt").DiscardSavedData
    Report.OpenSubreport("RPTOUTWARD.rpt").VerifyOnEveryPrint = True
    Set CRXFormulaFields = Report.OpenSubreport("RPTOUTWARD.rpt").FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "ITEM WISE INWARD OUTWARD MOVEMENT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            GRDOUTWARD.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    REPFLAG = True
    
    GRDOUTWARD.TextMatrix(0, 0) = "SL"
    GRDOUTWARD.TextMatrix(0, 1) = "TYPE"
    GRDOUTWARD.TextMatrix(0, 2) = "CUSTOMER"
    GRDOUTWARD.TextMatrix(0, 3) = "QTY"
    GRDOUTWARD.TextMatrix(0, 4) = "TAX" '"FREE"
    GRDOUTWARD.TextMatrix(0, 5) = "RATE" '"Pack"
    GRDOUTWARD.TextMatrix(0, 6) = "NET RATE"
    GRDOUTWARD.TextMatrix(0, 7) = "INV #"
    GRDOUTWARD.TextMatrix(0, 8) = "INV DATE"
    GRDOUTWARD.TextMatrix(0, 9) = "AMOUNT"

    GRDOUTWARD.ColWidth(0) = 400
    GRDOUTWARD.ColWidth(1) = 0
    GRDOUTWARD.ColWidth(2) = 3200
    GRDOUTWARD.ColWidth(3) = 1000
    GRDOUTWARD.ColWidth(4) = 1000
    GRDOUTWARD.ColWidth(5) = 1000
    GRDOUTWARD.ColWidth(6) = 900
    GRDOUTWARD.ColWidth(7) = 1000
    GRDOUTWARD.ColWidth(8) = 1200
    GRDOUTWARD.ColWidth(9) = 1600
    
    GRDOUTWARD.ColAlignment(0) = 4
    GRDOUTWARD.ColAlignment(1) = 1
    GRDOUTWARD.ColAlignment(2) = 1
    GRDOUTWARD.ColAlignment(3) = 1
    GRDOUTWARD.ColAlignment(4) = 1
    GRDOUTWARD.ColAlignment(5) = 4
    GRDOUTWARD.ColAlignment(6) = 1
    GRDOUTWARD.ColAlignment(7) = 4
    GRDOUTWARD.ColAlignment(8) = 4
    GRDOUTWARD.ColAlignment(9) = 4
    
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    'Me.Height = 9990
    'Me.Width = 18555
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If REPFLAG = False Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub OPTOUTCUST_Click()
    Call Fillgrid2
End Sub

Private Sub OPTOUTDATE_Click()
    Call Fillgrid2
End Sub

Private Sub OptSupplier_Click()
    Call Fillgrid2
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo ErrHand
    If REPFLAG = True Then
        RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ucase(CATEGORY) = 'SERVICE CHARGE' OR ucase(CATEGORY) = 'SERVICES') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND (ucase(CATEGORY) = 'SERVICE CHARGE' OR ucase(CATEGORY) = 'SERVICES') ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"

    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList2_Click()
    Call Fillgrid2
    ''''''''LBLBALANCE.Caption = Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption)
End Sub

Private Function Fillgrid2()
    Dim rststock As ADODB.Recordset
    Dim RSTTEMP As ADODB.Recordset
    Dim M As Integer
    Dim E_DATE As Date
    Dim i, Full_Qty, Loose_qty As Double
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ErrHand
    
    
    
    
    Screen.MousePointer = vbHourglass
    lblbalance.Caption = ""
    LBLOUTWARD.Caption = ""
    LblLoose.Caption = ""
    Label1(3).Caption = ""
    
    'db.Execute "delete From TEMPTRX"
    'Set RSTTEMP = New ADODB.Recordset
    'RSTTEMP.Open "SELECT *  FROM TEMPTRX", db, adOpenStatic, adLockOptimistic, adCmdText
    i = 0
    Full_Qty = 0
    Loose_qty = 0
    GRDOUTWARD.rows = 1
    Set rststock = New ADODB.Recordset
    If OPTOUTDATE.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE = 'RW' OR TRX_TYPE='SR' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
    If OPTOUTCUST.Value = True Then rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE = 'RW' OR TRX_TYPE='SR' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DESC ASC, VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
    'rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND CST <>2", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        i = i + 1
        GRDOUTWARD.rows = GRDOUTWARD.rows + 1
        GRDOUTWARD.FixedRows = 1
        GRDOUTWARD.TextMatrix(i, 0) = i
'        Select Case rststock!CST
'            Case 0
'               GRDOUTWARD.TextMatrix(i, 1) = "SALES"
'            Case 1
'               GRDOUTWARD.TextMatrix(i, 1) = "DELIVEREY"
'            Case 2
'               GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
'        End Select
        GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
        Select Case rststock!TRX_TYPE
            Case "SI"
                GRDOUTWARD.TextMatrix(i, 1) = "WHOLESALE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "RI"
                GRDOUTWARD.TextMatrix(i, 1) = "RETAIL"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "WO"
                GRDOUTWARD.TextMatrix(i, 1) = "WO"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DN"
                GRDOUTWARD.TextMatrix(i, 1) = "DELIVERY"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "PR"
                GRDOUTWARD.TextMatrix(i, 1) = "PURCHASE RETURN"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "DG", "DM"
                GRDOUTWARD.TextMatrix(i, 1) = "DAMAGED GOODS"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "SR", "RW"
                GRDOUTWARD.TextMatrix(i, 1) = "To Service"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "GF"
                GRDOUTWARD.TextMatrix(i, 1) = "SAMPLE"
                GRDOUTWARD.TextMatrix(i, 2) = Mid(rststock!VCH_DESC, 15)
            Case "MI"
                GRDOUTWARD.TextMatrix(i, 1) = "FACTORY"
                GRDOUTWARD.TextMatrix(i, 2) = "FACTORY"
        End Select
        
        GRDOUTWARD.Tag = ""
        Set RSTTEMP = New ADODB.Recordset
        RSTTEMP.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly
        If Not (RSTTEMP.EOF And RSTTEMP.BOF) Then
            GRDOUTWARD.Tag = IIf(IsNull(RSTTEMP!PACK_TYPE), "", RSTTEMP!PACK_TYPE)
            tXTMEDICINE.Tag = IIf(IsNull(RSTTEMP!LOOSE_PACK), "", RSTTEMP!LOOSE_PACK)
        End If
        RSTTEMP.Close
        Set RSTTEMP = Nothing
        
        GRDOUTWARD.TextMatrix(i, 3) = rststock!QTY
        GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!SALES_TAX), "", rststock!SALES_TAX)
        GRDOUTWARD.TextMatrix(i, 5) = IIf(IsNull(rststock!P_RETAILWOTAX), "", rststock!P_RETAILWOTAX)
        Full_Qty = Full_Qty + (Val(GRDOUTWARD.TextMatrix(i, 3)) + Val(GRDOUTWARD.TextMatrix(i, 4))) * GRDOUTWARD.TextMatrix(i, 5)
        
        'GRDOUTWARD.TextMatrix(i, 3) = rststock!QTY
        'GRDOUTWARD.TextMatrix(i, 4) = IIf(IsNull(rststock!FREE_QTY), "", rststock!FREE_QTY)
        GRDOUTWARD.TextMatrix(i, 6) = Format(rststock!P_RETAIL, "0.00")
        GRDOUTWARD.TextMatrix(i, 7) = rststock!TRX_TYPE & "-" & rststock!VCH_NO
        GRDOUTWARD.TextMatrix(i, 8) = Format(rststock!VCH_DATE, "dd/mm/yyyy")
        GRDOUTWARD.TextMatrix(i, 9) = IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL)
        lblbalance.Caption = Val(lblbalance.Caption) + Val(GRDOUTWARD.TextMatrix(i, 9))
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLOUTWARD.Caption = Full_Qty
    If Loose_qty > 0 Then
        LblLoose.Visible = True
        Label1(2).Visible = True
        Label1(3).Visible = True
        LblLoose.Caption = Loose_qty
        Label1(3).Caption = GRDOUTWARD.Tag
    Else
        LblLoose.Visible = False
        Label1(2).Visible = False
        Label1(3).Visible = False
        Label1(3).Caption = ""
        LblLoose.Caption = ""
    End If
    
    LBLHEAD(1).Caption = "OUTWARD DETAILS"
    Screen.MousePointer = vbNormal
    'If Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption) <> Val(LBLBALANCE.Caption) Then lblmanual.Visible = True Else lblmanual.Visible = False
    lblbalance.Caption = Format(lblbalance.Caption, "0.00")
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

