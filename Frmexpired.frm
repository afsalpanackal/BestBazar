VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMEXP 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPIRY"
   ClientHeight    =   8610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19800
   ControlBox      =   0   'False
   Icon            =   "Frmexpired.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   19800
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export to Excel"
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
      Left            =   6870
      TabIndex        =   15
      Top             =   75
      Width           =   1605
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
      Left            =   1440
      TabIndex        =   14
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox txtactqty 
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
      Left            =   5370
      TabIndex        =   11
      Top             =   8925
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   4725
      TabIndex        =   9
      Top             =   1605
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   9075
      _Version        =   393216
      Rows            =   1
      Cols            =   14
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      SelectionMode   =   1
      Appearance      =   0
      GridLineWidth   =   2
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
   Begin VB.Frame FRMEDISPLAY 
      BackColor       =   &H00C0C0FF&
      Height          =   630
      Left            =   15
      TabIndex        =   4
      Top             =   -90
      Width           =   6645
      Begin VB.PictureBox picChecked 
         Height          =   285
         Left            =   6015
         Picture         =   "Frmexpired.frx":000C
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   13
         Top             =   165
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picUnchecked 
         Height          =   285
         Left            =   6360
         Picture         =   "Frmexpired.frx":034E
         ScaleHeight     =   225
         ScaleWidth      =   225
         TabIndex        =   12
         Top             =   195
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.CheckBox CHKSELECT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   5355
         TabIndex        =   10
         Top             =   330
         Width           =   1215
      End
      Begin VB.TextBox TXTDAYS 
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
         Left            =   2085
         MaxLength       =   2
         TabIndex        =   7
         Top             =   225
         Width           =   600
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         Height          =   390
         Left            =   2790
         TabIndex        =   6
         Top             =   165
         Width           =   1230
      End
      Begin VB.CommandButton CMDPRINTEXPIRY 
         Caption         =   "&PRINT"
         Height          =   390
         Left            =   4080
         TabIndex        =   5
         Top             =   165
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER THE MONTH(s)"
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
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   1995
      End
   End
   Begin VB.OptionButton optbyname 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sort by Name"
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
      Left            =   3660
      TabIndex        =   2
      Top             =   8100
      Width           =   1830
   End
   Begin VB.OptionButton optbydate 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Sort by Exp. Date"
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
      Left            =   5640
      TabIndex        =   1
      Top             =   8100
      Value           =   -1  'True
      Width           =   2145
   End
   Begin VB.CommandButton CMDDELEXPIRY 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   60
      TabIndex        =   0
      Top             =   8100
      Width           =   1335
   End
   Begin MSFlexGridLib.MSFlexGrid grdEXPIRYLIST 
      Height          =   7365
      Left            =   15
      TabIndex        =   3
      Top             =   570
      Width           =   19770
      _ExtentX        =   34872
      _ExtentY        =   12991
      _Version        =   393216
      Rows            =   1
      Cols            =   17
      FixedRows       =   0
      RowHeightMin    =   450
      BackColor       =   16777215
      ForeColor       =   -2147483641
      BackColorFixed  =   8421504
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      Appearance      =   0
      GridLineWidth   =   2
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
End
Attribute VB_Name = "FRMEXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim strChecked As String

Private Sub CHKSELECT_Click()
    Dim i As Long
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    For i = 1 To grdEXPIRYLIST.rows - 1
        If CHKSELECT.Value = 1 Then
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 16) = "Y"
            End With
        Else
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 16) = "N"
            End With
        End If
    Next i
    Call fillcount
End Sub

Private Sub CMDDELEXPIRY_Click()
    Dim RSTEXP As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    
    Dim n As Long
    Dim M_DATA As Integer
    Dim K_DATA As Integer
    On Error GoTo ErrHand
    
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    If grdcount.rows = 0 Then
        MsgBox "NOTHING SELECTED!!!!", vbOKOnly, "DELETE !!!!"
        Exit Sub
    End If
    
    Dim RSTTRXFILE As ADODB.Recordset
    Dim MAXNO As Double
    
    MAXNO = 1
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select MAX(Val(VCH_NO)) From TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'EX'", db, adOpenForwardOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        MAXNO = IIf(IsNull(RSTTRXFILE.Fields(0)), 1, RSTTRXFILE.Fields(0) + 1)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'If MsgBox("ARE YOU SURE YOU WANT TO DELETE SL NO.  " & """" & grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 1) & """", vbYesNo, "EDIT....") = vbNo Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE ALL MARKED ITEMS", vbYesNo, "DELETE !!!") = vbNo Then Exit Sub
    For n = 0 To grdcount.rows - 1
        db.Execute ("Delete from EXPIRY where EXPIRY.EX_SLNO = '" & Val(grdcount.TextMatrix(n, 0)) & "'")
        M_DATA = 0
        K_DATA = 0
        Set RSTEXP = New ADODB.Recordset
        RSTEXP.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & grdcount.TextMatrix(n, 10) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        
        Do Until RSTEXP.EOF
            If Not (RSTEXP!LINE_NO = Val(grdcount.TextMatrix(n, 12)) And RSTEXP!VCH_NO = Val(grdcount.TextMatrix(n, 11))) Then GoTo SKIP
            
            RSTTRXFILE.AddNew
            RSTTRXFILE!TRX_TYPE = "EX"
            RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
            RSTTRXFILE!VCH_NO = MAXNO
            RSTTRXFILE!VCH_DATE = Format(Date, "DD/MM/YYYY")
            RSTTRXFILE!LINE_NO = n + 1
            RSTTRXFILE!Category = "GENERAL"
            RSTTRXFILE!LOOSE_FLAG = "L"
            RSTTRXFILE!ITEM_CODE = grdcount.TextMatrix(n, 10)
            RSTTRXFILE!ITEM_NAME = grdcount.TextMatrix(n, 1)
            RSTTRXFILE!QTY = Val(grdcount.TextMatrix(n, 4))
            RSTTRXFILE!ITEM_COST = Val(grdcount.TextMatrix(n, 5))
            RSTTRXFILE!MRP = Val(grdcount.TextMatrix(n, 5))
            RSTTRXFILE!PTR = Val(grdcount.TextMatrix(n, 5))
            RSTTRXFILE!SALES_PRICE = Val(grdcount.TextMatrix(n, 5))
            RSTTRXFILE!SALES_TAX = 0
            RSTTRXFILE!UNIT = Val(grdcount.TextMatrix(n, 3))
            RSTTRXFILE!PACK = ""
            RSTTRXFILE!VCH_DESC = "Issued to     Expiry"
            RSTTRXFILE!REF_NO = grdcount.TextMatrix(n, 6)
            RSTTRXFILE!ISSUE_QTY = 0
            RSTTRXFILE!CST = 0
            RSTTRXFILE!BAL_QTY = 0
            RSTTRXFILE!TRX_TOTAL = 0
            RSTTRXFILE!LINE_DISC = 0
            RSTTRXFILE!SCHEME = 0
            If IsDate(grdcount.TextMatrix(n, 2)) Then
                RSTTRXFILE!exp_date = grdcount.TextMatrix(n, 2)
            Else
                'RSTTRXFILE!EXP_DATE = Null
            End If
            RSTTRXFILE!FREE_QTY = 0
            RSTTRXFILE!CREATE_DATE = Date
            RSTTRXFILE!C_USER_ID = "SM"
            RSTTRXFILE!M_USER_ID = "311000"
            RSTTRXFILE.Update
            
            Set RSTITEM = New ADODB.Recordset
            RSTITEM.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdcount.TextMatrix(n, 10) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            
            M_DATA = Val(RSTITEM!ISSUE_QTY) + Val(grdcount.TextMatrix(n, 4))
            RSTITEM!ISSUE_QTY = M_DATA
            
            RSTITEM!ISSUE_VAL = 0
            
            K_DATA = RSTITEM!CLOSE_QTY - Val(grdcount.TextMatrix(n, 4))
            RSTITEM!CLOSE_QTY = K_DATA
            
            RSTITEM!CLOSE_VAL = 0
            RSTEXP!ISSUE_QTY = RSTEXP!BAL_QTY
            RSTEXP!BAL_QTY = 0
            RSTITEM.Update
            RSTEXP.Update
            RSTITEM.Close
            Set RSTITEM = Nothing
SKIP:
            RSTEXP.MoveNext
        Loop
        
        RSTEXP.Close
        
        Set RSTEXP = Nothing
    
    Next n
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Call FILLEXPIRYLIST
    grdcount.rows = 0
    Exit Sub
   
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CmDDisplay_Click()
    Dim RSTD As ADODB.Recordset
    Dim RSTE As ADODB.Recordset
    Dim RSTF As ADODB.Recordset
    
    Dim M_DATE As Date
    Dim E_DATE As Date
    
    Dim i As Long
    
    
    If FRMEXP.Txtdays.text = "" Then Exit Sub
    
    On Error GoTo ErrHand
    
    i = LastDayOfMonth(Date)
    M_DATE = i & "/" & Month(Date) & "/" & Year(Date)
    
    E_DATE = DateAdd("m", Val(Txtdays.text), M_DATE)
    
    

    Screen.MousePointer = vbHourglass
    db.Execute ("DELETE FROM EXPIRY")
    
    On Error Resume Next
    db.Execute "Update rtrxfile set exp_date = Null WHERE exp_date = '0000-00-00'"
    db.Execute "Update rtrxfile set expiry = STR_TO_DATE(exp_date,'%d/%c/%Y') WHERE BAL_QTY > 0"
    err.Clear
    On Error GoTo ErrHand
    i = 0
    Set RSTD = New ADODB.Recordset
    RSTD.Open "SELECT * From EXPIRY", db, adOpenStatic, adLockOptimistic, adCmdText
    Set RSTE = New ADODB.Recordset
    'RSTE.Open "SELECT * from RTRXFILE WHERE BAL_QTY > 0 AND EXP_DATE <= '" & E_DATE & "' ORDER BY RTRXFILE.EXP_DATE", db, adOpenForwardOnly
    RSTE.Open "SELECT * from RTRXFILE WHERE BAL_QTY > 0 AND expiry <= '" & Format(E_DATE, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
    Do Until RSTE.EOF
        
        RSTD.AddNew
        i = i + 1
        RSTD!EX_SLNO = i
        RSTD!EX_ITEM = RSTE!ITEM_NAME
        RSTD!EX_PUR_INV = RSTE!PINV
        RSTD!EX_PUR_DATE = RSTE!VCH_DATE
        
        Set RSTF = New ADODB.Recordset
        RSTF.Open "SELECT * From ITEMMAST WHERE ITEM_CODE ='" & RSTE!ITEM_CODE & "'", db, adOpenForwardOnly
        If Not (RSTF.EOF And RSTF.BOF) Then
            RSTD!EX_MFGR = RSTF!MANUFACTURER
        End If
        
        If Len(RSTE!VCH_DESC) > 14 Then
            If Mid(RSTE!VCH_DESC, 1, 14) = "Received From " Then
                RSTD!EX_DISTI = Mid(RSTE!VCH_DESC, 15)
            Else
                RSTD!EX_DISTI = RSTE!VCH_DESC
            End If
        Else
            RSTD!EX_DISTI = RSTE!VCH_DESC
        End If
        RSTD!EX_BATCH = RSTE!REF_NO
        RSTD!EX_DATE = Format(RSTE!exp_date, "DD/MM/YY")
        RSTD!EX_QTY = RSTE!BAL_QTY
        RSTD!EX_MRP = RSTE!MRP
        RSTD!EX_SETTLE = ""
        RSTD!EX_VALUE = Val(RSTE!MRP) * Val(RSTE!BAL_QTY)
        RSTD!VCH_NO = RSTE!VCH_NO
        RSTD!LINE_NO = RSTE!LINE_NO
        RSTD!ITEM_CODE = RSTE!ITEM_CODE
        RSTD!EX_UNIT = RSTE!UNIT
        RSTD!PINV = RSTE!PINV
        
        RSTD.Update
        
        RSTF.Close
        Set RSTF = Nothing
            
        RSTE.MoveNext
            
    Loop

    RSTE.Close
 
     Set RSTE = Nothing

    
    RSTD.Close
    Set RSTD = Nothing
    
    Call FILLEXPIRYLIST
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub


Private Sub CmdExport_Click()
    If frmLogin.rs!Level <> "0" Then MsgBox "Permission Denied", vbOKOnly, "Price Analysis"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i As Long
    Dim n As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    'Create an Excel instalce.
    Set oApp = CreateObject("Excel.Application")
    Set oWB = oApp.Workbooks.Add
    Set oWS = oWB.Worksheets(1)
    

    
    
'    xlRange = oWS.Range("A1", "C1")
'    xlRange.Font.Bold = True
'    xlRange.ColumnWidth = 15
'    'xlRange.Value = {"First Name", "Last Name", "Last Service"}
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).LineStyle = Excel.XlLineStyle.xlContinuous
'    xlRange.Borders(Excel.XlBordersIndex.xlEdgeBottom).Weight = 2
'
'    xlRange = oWS.Range("C1", "C999")
'    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter
'    xlRange.ColumnWidth = 12
    
    'If Sum_flag = False Then
        oWS.Range("A1", "J1").Merge
        oWS.Range("A1", "J1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "J2").Merge
        oWS.Range("A2", "J2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    oWS.Range("F:F").ColumnWidth = 12
    oWS.Range("G:G").ColumnWidth = 12
    oWS.Range("H:H").ColumnWidth = 12
    oWS.Range("I:I").ColumnWidth = 12
    oWS.Range("J:J").ColumnWidth = 12
    oWS.Range("K:K").ColumnWidth = 12
    oWS.Range("L:L").ColumnWidth = 12
    oWS.Range("M:M").ColumnWidth = 12
    oWS.Range("N:N").ColumnWidth = 12
    oWS.Range("O:O").ColumnWidth = 12
    oWS.Range("P:P").ColumnWidth = 12
    oWS.Range("Q:Q").ColumnWidth = 12
    oWS.Range("R:R").ColumnWidth = 12
    oWS.Range("S:S").ColumnWidth = 12
    oWS.Range("T:T").ColumnWidth = 12
    oWS.Range("U:U").ColumnWidth = 12
    oWS.Range("V:V").ColumnWidth = 12
    
    oWS.Range("A1").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 14            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    'oApp.Columns("A:A").EntireColumn.AutoFit     '-- autofitted column

    oWS.Range("A2").Select                      '-- particular cell selection
    oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 11            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True

'    Range("C2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("C:C").EntireColumn.AutoFit     '-- autofitted column
'
'
'    Range("D2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("D:D").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("E2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("E:E").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("F2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("F:F").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("G2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("G:G").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("H2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("H:H").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("I2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("I:I").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("J2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("J:J").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("K2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("K:K").EntireColumn.AutoFit     '-- autofitted column
'
'    Range("L2").Select                      '-- particular cell selection
'    ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
'    Selection.Font.Bold = True              '-- enabled bold cell style
'    Columns("L:L").EntireColumn.AutoFit     '-- autofitted column

'    oWB.ActiveSheet.Font.Name = "Arial"
'    oApp.ActiveSheet.Font.Name = "Arial"
'    oWB.Font.Size = "11"
'    oWB.Font.Bold = True
    oWS.Range("A" & 1).Value = MDIMAIN.StatusBar.Panels(5).text
    oWS.Range("A" & 2).Value = "EXPIRY REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 1)
    oWS.Range("B" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 3)
    oWS.Range("C" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 4)
    oWS.Range("D" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 5)
    oWS.Range("E" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 6)
    oWS.Range("F" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 7)
    oWS.Range("G" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 8)
    oWS.Range("H" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 9)
    oWS.Range("I" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 10)
    oWS.Range("J" & 3).Value = grdEXPIRYLIST.TextMatrix(0, 11)
    
    On Error GoTo ErrHand
    
    i = 4
    For n = 1 To grdEXPIRYLIST.rows - 1
        oWS.Range("A" & i).Value = grdEXPIRYLIST.TextMatrix(n, 1)
        oWS.Range("B" & i).Value = grdEXPIRYLIST.TextMatrix(n, 3)
        oWS.Range("C" & i).Value = "'" & grdEXPIRYLIST.TextMatrix(n, 4)
        oWS.Range("D" & i).Value = grdEXPIRYLIST.TextMatrix(n, 5)
        oWS.Range("E" & i).Value = grdEXPIRYLIST.TextMatrix(n, 6)
        oWS.Range("F" & i).Value = grdEXPIRYLIST.TextMatrix(n, 7)
        oWS.Range("G" & i).Value = grdEXPIRYLIST.TextMatrix(n, 8)
        oWS.Range("H" & i).Value = grdEXPIRYLIST.TextMatrix(n, 9)
        oWS.Range("I" & i).Value = grdEXPIRYLIST.TextMatrix(n, 10)
        oWS.Range("J" & i).Value = grdEXPIRYLIST.TextMatrix(n, 11)
        On Error GoTo ErrHand
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    'oApp.ActiveCell.FormulaR1C1 = "123"          '-- cell text fill
    oApp.Selection.HorizontalAlignment = xlRight
    oApp.Selection.Font.Name = "Arial"             '-- enabled bold cell style
    oApp.Selection.Font.Size = 13            '-- enabled bold cell style
    oApp.Selection.Font.Bold = True
    
   
SKIP:
    oApp.Visible = True
    
    
'    Set oWB = Nothing
'    oApp.Quit
'    Set oApp = Nothing
'
    
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox err.Description
End Sub

Private Sub CMDPRINTEXPIRY_Click()
    Dim n As Long
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    On Error GoTo ErrHand
    db.Execute "DELETE FROM TEMPSTK"
    Screen.MousePointer = vbHourglass
    Dim RSTTEM As ADODB.Recordset
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK", db, adOpenStatic, adLockOptimistic, adCmdText
    For n = 0 To grdcount.rows - 1
        RSTTEM.AddNew
        'RSTTEM!ITEM_CODE = GRDSTOCK.TextMatrix(i, 1)
        RSTTEM!ITEM_NAME = grdcount.TextMatrix(n, 1)
        RSTTEM!INQTY = grdcount.TextMatrix(n, 4)
        RSTTEM!DIFF_QTY = grdcount.TextMatrix(n, 2)
        'RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 5)
        'RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 6)
        'RSTTEM!DIFF_QTY = Trim(GRDSTOCK.TextMatrix(i, 7))
        'RSTTEM!OPQTY = 0 'GRDSTOCK.TextMatrix(i, 3)
        'RSTTEM!OPVAL = 0 'GRDSTOCK.TextMatrix(i, 4)
        'RSTTEM!INQTY_VAL = GRDSTOCK.TextMatrix(i, 4)
'        RSTTEM!OUTQTY_VAL = GRDSTOCK.TextMatrix(i, 6)
'        RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 7)
'        RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 8)
        
        'RSTTEM!DIFF_QTY = 0 'GRDSTOCK.TextMatrix(i, 11)
        'RSTTEM!DIFF_VAL = 0 'GRDSTOCK.TextMatrix(i, 11)
        RSTTEM.Update
    Next n
    RSTTEM.Close
    Set RSTTEM = Nothing
    
    frmreport.Caption = "EXPIRY REPORT"
    ReportNameVar = "D:\EzBiz\RptEXPReport"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
    
    For n = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(n).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(n).Name & " ")
            Report.Database.SetDataSource oRs, 3, n
            Set oRs = Nothing
        End If
    Next n
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next

    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    'Width = 16110
    'Height = 9765
    Left = 0
    Top = 0
    grdcount.rows = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub grdEXPIRYLIST_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If grdEXPIRYLIST.rows = 1 Then Exit Sub
    With grdEXPIRYLIST
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If grdEXPIRYLIST.Col = 0 Then
                If grdEXPIRYLIST.CellPicture = picChecked Then
                    Set grdEXPIRYLIST.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 16) = "Y"
                    Call fillcount
                Else
                    Set grdEXPIRYLIST.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 16) = "N"
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Sub grdEXPIRYLIST_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
            Call grdEXPIRYLIST_Click
    End Select
End Sub

Private Sub optbydate_Click()
    Call FILLEXPIRYLIST
End Sub

Private Sub optbyname_Click()
    Call FILLEXPIRYLIST
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub Txtdays_GotFocus()
    Txtdays.SelStart = 0
    Txtdays.SelLength = Len(Txtdays.text)
End Sub

Private Sub Txtdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtdays.text) = 0 Then Exit Sub
            CMDDISPLAY.SetFocus
                        
    End Select
End Sub

Private Sub Txtdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, 45
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub FILLEXPIRYLIST()
    Dim i As Long
    Dim rstexplist As ADODB.Recordset
    
    grdEXPIRYLIST.Visible = False
    Screen.MousePointer = vbHourglass
    i = 0
    grdEXPIRYLIST.TextMatrix(0, 0) = ""
    grdEXPIRYLIST.TextMatrix(0, 1) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 2) = "SL CODE"
    grdEXPIRYLIST.TextMatrix(0, 3) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 4) = "EXPIRY"
    grdEXPIRYLIST.TextMatrix(0, 5) = "PACK"
    grdEXPIRYLIST.TextMatrix(0, 6) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 7) = "MRP"
    grdEXPIRYLIST.TextMatrix(0, 8) = "BATCH"
    grdEXPIRYLIST.TextMatrix(0, 9) = "MFGR"
    grdEXPIRYLIST.TextMatrix(0, 10) = "SUPPLIER"
    grdEXPIRYLIST.TextMatrix(0, 11) = "INVOICE NO"
    grdEXPIRYLIST.TextMatrix(0, 12) = "ITEM CODE"
    grdEXPIRYLIST.TextMatrix(0, 13) = "VCH NO"
    grdEXPIRYLIST.TextMatrix(0, 14) = "LINE NO"
    grdEXPIRYLIST.TextMatrix(0, 15) = "INV DATE"
    grdEXPIRYLIST.TextMatrix(0, 16) = "FLAG"
    
    grdEXPIRYLIST.ColWidth(0) = 300
    grdEXPIRYLIST.ColWidth(1) = 500
    grdEXPIRYLIST.ColWidth(2) = 0
    grdEXPIRYLIST.ColWidth(3) = 6000
    grdEXPIRYLIST.ColWidth(4) = 1000
    grdEXPIRYLIST.ColWidth(5) = 650
    grdEXPIRYLIST.ColWidth(6) = 650
    grdEXPIRYLIST.ColWidth(7) = 900
    grdEXPIRYLIST.ColWidth(8) = 1300
    grdEXPIRYLIST.ColWidth(9) = 900
    grdEXPIRYLIST.ColWidth(10) = 2700
    grdEXPIRYLIST.ColWidth(11) = 1500
    grdEXPIRYLIST.ColWidth(12) = 0
    grdEXPIRYLIST.ColWidth(13) = 0
    grdEXPIRYLIST.ColWidth(14) = 0
    grdEXPIRYLIST.ColWidth(15) = 1300
    grdEXPIRYLIST.ColWidth(16) = 0
    
    grdEXPIRYLIST.ColAlignment(0) = 4
    grdEXPIRYLIST.ColAlignment(0) = 4
    grdEXPIRYLIST.ColAlignment(3) = 4
    grdEXPIRYLIST.ColAlignment(4) = 4
    grdEXPIRYLIST.ColAlignment(5) = 4
    grdEXPIRYLIST.ColAlignment(6) = 4
    grdEXPIRYLIST.ColAlignment(7) = 4
    grdEXPIRYLIST.ColAlignment(8) = 4
    grdEXPIRYLIST.ColAlignment(9) = 4
    grdEXPIRYLIST.ColAlignment(10) = 4
    grdEXPIRYLIST.ColAlignment(11) = 4
    grdEXPIRYLIST.ColAlignment(15) = 4
    
    On Error GoTo ErrHand
    i = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.rows = 1
    Set rstexplist = New ADODB.Recordset
    If optbydate.Value = True Then
        rstexplist.Open "select * from EXPIRY ORDER BY STR_TO_DATE(EX_DATE,'%d/%c/%Y')", db, adOpenForwardOnly
    Else
        rstexplist.Open "select * from EXPIRY ORDER BY EX_ITEM", db, adOpenForwardOnly
    End If
    Do Until rstexplist.EOF
        i = i + 1
        grdEXPIRYLIST.rows = grdEXPIRYLIST.rows + 1
        grdEXPIRYLIST.FixedRows = 1
        'grdEXPIRYLIST.TextMatrix(i, 0) = i
        grdEXPIRYLIST.TextMatrix(i, 1) = i
        grdEXPIRYLIST.TextMatrix(i, 2) = rstexplist!EX_SLNO
        grdEXPIRYLIST.TextMatrix(i, 3) = rstexplist!EX_ITEM
        grdEXPIRYLIST.TextMatrix(i, 4) = IIf(IsNull(rstexplist!EX_DATE), "", Format(rstexplist!EX_DATE, "MM/YY"))
        grdEXPIRYLIST.TextMatrix(i, 5) = IIf(IsNull(rstexplist!EX_UNIT), "", rstexplist!EX_UNIT)
        grdEXPIRYLIST.TextMatrix(i, 6) = IIf(IsNull(rstexplist!EX_QTY), "", rstexplist!EX_QTY)
        grdEXPIRYLIST.TextMatrix(i, 7) = IIf(IsNull(rstexplist!EX_MRP), "", Format(rstexplist!EX_MRP, ".000"))
        grdEXPIRYLIST.TextMatrix(i, 8) = IIf(IsNull(rstexplist!EX_BATCH), "", rstexplist!EX_BATCH)
        grdEXPIRYLIST.TextMatrix(i, 9) = IIf(IsNull(rstexplist!EX_MFGR), "", rstexplist!EX_MFGR)
        grdEXPIRYLIST.TextMatrix(i, 10) = IIf(IsNull(rstexplist!EX_DISTI), "", rstexplist!EX_DISTI)
        grdEXPIRYLIST.TextMatrix(i, 11) = IIf(IsNull(rstexplist!PINV), "", rstexplist!PINV)
        grdEXPIRYLIST.TextMatrix(i, 12) = IIf(IsNull(rstexplist!ITEM_CODE), "", rstexplist!ITEM_CODE)
        grdEXPIRYLIST.TextMatrix(i, 13) = IIf(IsNull(rstexplist!VCH_NO), "", rstexplist!VCH_NO)
        grdEXPIRYLIST.TextMatrix(i, 14) = IIf(IsNull(rstexplist!LINE_NO), "", rstexplist!LINE_NO)
        grdEXPIRYLIST.TextMatrix(i, 15) = IIf(IsNull(rstexplist!EX_PUR_DATE), "", Format(rstexplist!EX_PUR_DATE, "DD/MM/YYYY"))
        grdEXPIRYLIST.TextMatrix(i, 16) = "N"
        With grdEXPIRYLIST
          .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
          Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
          .TextMatrix(i, 1) = i
        End With
        rstexplist.MoveNext
    Loop
    rstexplist.Close
    Set rstexplist = Nothing
    grdEXPIRYLIST.Visible = True
    CHKSELECT.Value = 0
    grdcount.rows = 0
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub EXPIRYReport()

    Dim n As Long
    db.Execute "DELETE FROM TEMPSTK"
    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
    '//NOTE : Report file name should never contain blank space.
    Open "D:\EzBiz\Report.PRN" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    Print #1, Space(3) & AlignLeft(" SL", 2) & Space(1) & _
            AlignLeft("ITEM NAME", 11) & Space(12) & _
            AlignLeft("EXP DATE", 12) & _
            AlignLeft("QTY", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
            
    Print #1, Space(7) & RepeatString("-", 118)
    
    For n = 0 To grdcount.rows - 1
        Print #1, Space(2) & AlignRight(str(n + 1), 3) & Space(2) & _
                AlignLeft(grdcount.TextMatrix(n, 1), 20) & Space(5) & _
                AlignLeft(grdcount.TextMatrix(n, 2), 7) & Space(2) & _
                AlignRight(grdcount.TextMatrix(n, 4), 4) & Space(1) & _
                Chr(27) & Chr(72)  '//Bold Ends
    Next n
    
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    Close #1 '//Closing the file
    
End Sub

Private Function fillcount()
    Dim i, n As Long
    
    grdcount.rows = 0
    i = 0
    On Error GoTo ErrHand
    For n = 1 To grdEXPIRYLIST.rows - 1
        If grdEXPIRYLIST.TextMatrix(n, 16) = "Y" Then
            grdcount.rows = grdcount.rows + 1
            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(n, 2)
            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(n, 3)
            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(n, 4)
            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(n, 5)
            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(n, 6)
            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(n, 7)
            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(n, 8)
            grdcount.TextMatrix(i, 7) = grdEXPIRYLIST.TextMatrix(n, 9)
            grdcount.TextMatrix(i, 8) = grdEXPIRYLIST.TextMatrix(n, 10)
            grdcount.TextMatrix(i, 9) = grdEXPIRYLIST.TextMatrix(n, 11)
            grdcount.TextMatrix(i, 10) = grdEXPIRYLIST.TextMatrix(n, 12)
            grdcount.TextMatrix(i, 11) = grdEXPIRYLIST.TextMatrix(n, 13)
            grdcount.TextMatrix(i, 12) = grdEXPIRYLIST.TextMatrix(n, 14)
            grdcount.TextMatrix(i, 13) = grdEXPIRYLIST.TextMatrix(n, 15)
            i = i + 1
        End If
    Next n
    Exit Function
ErrHand:
    MsgBox err.Description
    
End Function

'Private Function markitems()
'    Dim i, n As Long
'
'    i = 0
'    On Error GoTo ErrHand
'    For n = 1 To grdEXPIRYLIST.rows - 1
'        If grdEXPIRYLIST.TextMatrix(n, 16) = "Y" Then
'            grdcount.rows = grdcount.rows + 1
'            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(n, 2)
'            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(n, 3)
'            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(n, 4)
'            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(n, 5)
'            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(n, 6)
'            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(n, 7)
'            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(n, 8)
'            grdcount.TextMatrix(i, 7) = grdEXPIRYLIST.TextMatrix(n, 9)
'            grdcount.TextMatrix(i, 8) = grdEXPIRYLIST.TextMatrix(n, 10)
'            grdcount.TextMatrix(i, 9) = grdEXPIRYLIST.TextMatrix(n, 11)
'            grdcount.TextMatrix(i, 10) = grdEXPIRYLIST.TextMatrix(n, 12)
'            grdcount.TextMatrix(i, 11) = grdEXPIRYLIST.TextMatrix(n, 13)
'            grdcount.TextMatrix(i, 12) = grdEXPIRYLIST.TextMatrix(n, 14)
'            grdcount.TextMatrix(i, 13) = grdEXPIRYLIST.TextMatrix(n, 15)
'            i = i + 1
'        End If
'    Next n
'    Exit Function
'ErrHand:
'    MsgBox err.Description
'
'End Function
