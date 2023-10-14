VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmAMC 
   BackColor       =   &H00CFEFE4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AMC Reminder"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17655
   Icon            =   "frmAMC.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   17655
   Begin VB.CommandButton CmdExport 
      Caption         =   "&Export"
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
      Left            =   45
      TabIndex        =   13
      Top             =   8190
      Width           =   1215
   End
   Begin VB.CommandButton CMDDETAILS 
      Caption         =   "&PRINT"
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
      Left            =   1290
      TabIndex        =   4
      Top             =   8190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFEFE4&
      Height          =   660
      Left            =   0
      TabIndex        =   2
      Top             =   -75
      Width           =   14835
      Begin VB.OptionButton OptAMC 
         BackColor       =   &H00CFEFE4&
         Caption         =   "AMC - Already done"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   6630
         MaskColor       =   &H00CFEFE4&
         TabIndex        =   12
         Top             =   255
         Width           =   2145
      End
      Begin VB.OptionButton Optnoamc 
         BackColor       =   &H00CFEFE4&
         Caption         =   "AMC - Not done yet"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   4245
         MaskColor       =   &H00CFEFE4&
         TabIndex        =   11
         Top             =   240
         Width           =   2145
      End
      Begin VB.OptionButton OPTDUE 
         BackColor       =   &H00CFEFE4&
         Caption         =   "AMC due within"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   210
         MaskColor       =   &H00CFEFE4&
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1830
      End
      Begin VB.TextBox Txtdays 
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
         ForeColor       =   &H000000FF&
         Height          =   390
         Left            =   2055
         MaxLength       =   2
         TabIndex        =   5
         Top             =   180
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   10515
         TabIndex        =   6
         Top             =   45
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   192
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   103940099
         CurrentDate     =   40498
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CFEFE4&
         Caption         =   "days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   3180
         TabIndex        =   3
         Top             =   240
         Width           =   1050
      End
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
      Height          =   450
      Left            =   13710
      TabIndex        =   1
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display"
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
      Left            =   12315
      TabIndex        =   0
      Top             =   8190
      Width           =   1200
   End
   Begin VB.Frame Frame2 
      Height          =   7695
      Left            =   15
      TabIndex        =   7
      Top             =   495
      Width           =   17610
      Begin VB.ComboBox CMBCHANGE 
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
         ForeColor       =   &H000000FF&
         Height          =   360
         ItemData        =   "frmAMC.frx":08CA
         Left            =   9480
         List            =   "frmAMC.frx":08D4
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   3390
         Visible         =   0   'False
         Width           =   1200
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   7560
         Left            =   0
         TabIndex        =   8
         Top             =   105
         Width           =   17580
         _ExtentX        =   31009
         _ExtentY        =   13335
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "FrmAMC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY_REC As New ADODB.Recordset
Dim PHY_FLAG As Boolean

Private Sub CMBCHANGE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo Errhand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 7  'AMC DONE
                    If CMBCHANGE.ListIndex = -1 Then CMBCHANGE.ListIndex = 0
                    Select Case CMBCHANGE.ListIndex
                        Case 0
                            db.Execute "Update TRXMAST set AMC_FLAG = 'Y' where VCH_NO = " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9) & "' AND TRX_YEAR = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10) & "'"
                        Case 1
                            db.Execute "Update TRXMAST set AMC_FLAG = 'N' where VCH_NO = " & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 9) & "' AND TRX_YEAR = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10) & "'"
                    End Select
                    GRDSTOCK.Enabled = True
                    CMBCHANGE.Visible = False
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = CMBCHANGE.Text
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            CMBCHANGE.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDDETAILS_Click()
    Dim RSTTEM As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
    db.Execute "DELETE * FROM TEMPSTK"
    
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "SOLD"
    GRDSTOCK.TextMatrix(0, 4) = "SHELF"
    GRDSTOCK.TextMatrix(0, 5) = "RQD QTY"
    GRDSTOCK.TextMatrix(0, 6) = "MRP"
    GRDSTOCK.TextMatrix(0, 7) = "SUPPLIER"
    
    If GRDSTOCK.Rows <= 1 Then Exit Sub
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To GRDSTOCK.Rows - 1
        RSTTEM.AddNew
        RSTTEM!ITEM_CODE = GRDSTOCK.TextMatrix(i, 1)
        RSTTEM!ITEM_NAME = GRDSTOCK.TextMatrix(i, 2)
        RSTTEM!INQTY = GRDSTOCK.TextMatrix(i, 3)
        RSTTEM!OUTQTY = GRDSTOCK.TextMatrix(i, 4)
        RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 5)
        RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 6)
        RSTTEM!DIFF_QTY = Trim(GRDSTOCK.TextMatrix(i, 7))
        'RSTTEM!OPQTY = 0 'GRDSTOCK.TextMatrix(i, 3)
        'RSTTEM!OPVAL = 0 'GRDSTOCK.TextMatrix(i, 4)
        'RSTTEM!INQTY_VAL = GRDSTOCK.TextMatrix(i, 4)
'        RSTTEM!OUTQTY_VAL = GRDSTOCK.TextMatrix(i, 6)
'        RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 7)
'        RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 8)
        
        'RSTTEM!DIFF_QTY = 0 'GRDSTOCK.TextMatrix(i, 11)
        'RSTTEM!DIFF_VAL = 0 'GRDSTOCK.TextMatrix(i, 11)
        RSTTEM.Update
    Next i
    RSTTEM.Close
    Set RSTTEM = Nothing
    
    frmreport.Caption = "STOCK REPORT"
    ReportNameVar = "D:\EzBiz\RptReport.RPT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "C:\Users\Public\WINSYS.SYS", "admin", "###DATABASE%%%RET"
    Next i
    Report.OpenSubreport("RptReport.RPT").DiscardSavedData
    Report.OpenSubreport("RptReport.RPT").VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'MY SHOP, ALAPPUZHA'"
    Next

    Call GENERATEREPORT
    
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Long
    Dim DDate As Date
    On Error GoTo Errhand
    
    If OPTDUE.value = True Then
        If Val(Txtdays.Text) = 0 Then
            MsgBox "Please enter the no. of days", vbOKOnly, "Reminder"
            Txtdays.SetFocus
            Exit Sub
        End If
        DTFROM.value = DateAdd("d", -(365 - Val(Txtdays.Text)), Date)
    End If
    
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    
    Dim RSTCOMPANY As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND [DOM] <=# " & Format(DTFROM.value, "MM,DD") & " # AND [DOM] >=# " & Format(Date, "MM,DD") & " # ORDER BY ACT_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    If OPTDUE.value = True Then
        rststock.Open "SELECT * From TRXMAST WHERE (isnull(AMC_FLAG) or AMC_FLAG <> 'Y') AND VCH_DATE <= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO') ORDER BY VCH_DATE, VCH_NO ", db, adOpenStatic, adLockReadOnly
    ElseIf OptAMC.value = True Then
        rststock.Open "SELECT * From TRXMAST WHERE AMC_FLAG = 'Y' AND VCH_DATE <= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO') ORDER BY VCH_DATE, VCH_NO ", db, adOpenStatic, adLockReadOnly
    Else
        rststock.Open "SELECT * From TRXMAST WHERE (isnull(AMC_FLAG) or AMC_FLAG <> 'Y') AND (TRX_TYPE='SV' OR TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='HI' OR TRX_TYPE='RI' OR TRX_TYPE='WO') ORDER BY VCH_DATE, VCH_NO ", db, adOpenStatic, adLockReadOnly
    End If
    Do Until rststock.EOF
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!VCH_NO
        Set RSTCOMPANY = New ADODB.Recordset
        RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
            Select Case rststock!TRX_TYPE
                Case "HI"
                    GRDSTOCK.TextMatrix(i, 1) = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V) & rststock!VCH_NO & IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
                Case "GI"
                    GRDSTOCK.TextMatrix(i, 1) = IIf(IsNull(RSTCOMPANY!PREFIX_8), "", RSTCOMPANY!PREFIX_8V) & rststock!VCH_NO & IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8)
                Case "SV"
                    GRDSTOCK.TextMatrix(i, 1) = IIf(IsNull(RSTCOMPANY!PREFIX_8B), "", RSTCOMPANY!PREFIX_8V) & rststock!VCH_NO & IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8B)
                Case "WO"
                    GRDSTOCK.TextMatrix(i, 1) = "PT-" & rststock!VCH_NO
                Case Else
                    GRDSTOCK.TextMatrix(i, 1) = rststock!VCH_NO
            End Select
        End If
        RSTCOMPANY.Close
        Set RSTCOMPANY = Nothing
        GRDSTOCK.TextMatrix(i, 2) = Format(rststock!VCH_DATE, "DD/MMM/YYYY")
        GRDSTOCK.TextMatrix(i, 3) = IIf(IsNull(rststock!ACT_NAME), "", rststock!ACT_NAME)
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!BILL_ADDRESS), "", rststock!BILL_ADDRESS)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!PHONE), "", rststock!PHONE)
        GRDSTOCK.TextMatrix(i, 6) = 365 + DateDiff("d", Date, rststock!VCH_DATE)
        If Val(GRDSTOCK.TextMatrix(i, 6)) < 0 Then
            GRDSTOCK.Col = 6
            GRDSTOCK.Row = i
            GRDSTOCK.CellForeColor = vbRed
        End If
        Select Case rststock!AMC_FLAG
            Case "Y"
                GRDSTOCK.TextMatrix(i, 7) = "Yes"
            Case Else
                GRDSTOCK.TextMatrix(i, 7) = "No"
        End Select
        GRDSTOCK.TextMatrix(i, 8) = rststock!VCH_NO
        GRDSTOCK.TextMatrix(i, 9) = rststock!TRX_TYPE
        GRDSTOCK.TextMatrix(i, 10) = rststock!TRX_YEAR
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing

    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "Annual Service Contract Details"
    If MsgBox("Are you sure?", vbYesNo + vbDefaultButton2, "Stock Report") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, N As Long
    
    On Error GoTo Errhand
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
        oWS.Range("A1", "E1").Merge
        oWS.Range("A1", "E1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "E2").Merge
        oWS.Range("A2", "E2").HorizontalAlignment = xlCenter
    'End If
    oWS.Range("A:A").ColumnWidth = 6
    oWS.Range("B:B").ColumnWidth = 10
    oWS.Range("C:C").ColumnWidth = 12
    oWS.Range("D:D").ColumnWidth = 12
    oWS.Range("E:E").ColumnWidth = 12
    
    
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

    oWS.Range("A" & 1).value = MDIMAIN.StatusBar.Panels(5).Text
    oWS.Range("A" & 2).value = "ZERO STOCK REPORT"
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).value = GRDSTOCK.TextMatrix(0, 0)
    oWS.Range("B" & 3).value = GRDSTOCK.TextMatrix(0, 1)
    oWS.Range("C" & 3).value = GRDSTOCK.TextMatrix(0, 2)
    oWS.Range("D" & 3).value = GRDSTOCK.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).value = GRDSTOCK.TextMatrix(0, 4)
    oWS.Range("F" & 3).value = GRDSTOCK.TextMatrix(0, 5)
    oWS.Range("G" & 3).value = GRDSTOCK.TextMatrix(0, 6)
    oWS.Range("H" & 3).value = GRDSTOCK.TextMatrix(0, 7)
    
    On Error GoTo Errhand
    
    i = 4
    For N = 1 To GRDSTOCK.Rows - 1
        oWS.Range("A" & i).value = GRDSTOCK.TextMatrix(N, 0)
        oWS.Range("B" & i).value = GRDSTOCK.TextMatrix(N, 1)
        oWS.Range("C" & i).value = GRDSTOCK.TextMatrix(N, 2)
        oWS.Range("D" & i).value = GRDSTOCK.TextMatrix(N, 3)
        oWS.Range("E" & i).value = GRDSTOCK.TextMatrix(N, 4)
        oWS.Range("F" & i).value = GRDSTOCK.TextMatrix(N, 5)
        oWS.Range("G" & i).value = GRDSTOCK.TextMatrix(N, 6)
        oWS.Range("H" & i).value = GRDSTOCK.TextMatrix(N, 7)
        
        On Error GoTo Errhand
        i = i + 1
    Next N
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
Errhand:
    'On Error Resume Next
    Screen.MousePointer = vbNormal
    Set oWB = Nothing
    'oApp.Quit
    'Set oApp = Nothing
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "Bill No"
    GRDSTOCK.TextMatrix(0, 2) = "Bill Date"
    GRDSTOCK.TextMatrix(0, 3) = "CUSTOMER"
    GRDSTOCK.TextMatrix(0, 4) = "Address"
    GRDSTOCK.TextMatrix(0, 5) = "Phone"
    GRDSTOCK.TextMatrix(0, 6) = "Days Left"
    GRDSTOCK.TextMatrix(0, 7) = "AMC Done?"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 1200
    GRDSTOCK.ColWidth(2) = 1100
    GRDSTOCK.ColWidth(3) = 3000
    GRDSTOCK.ColWidth(4) = 6000
    GRDSTOCK.ColWidth(5) = 1300
    GRDSTOCK.ColWidth(6) = 1300
    GRDSTOCK.ColWidth(7) = 1300
    GRDSTOCK.ColWidth(8) = 0
    GRDSTOCK.ColWidth(9) = 0
    GRDSTOCK.ColWidth(10) = 0
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 4
    GRDSTOCK.ColAlignment(2) = 4
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 1
    GRDSTOCK.ColAlignment(5) = 1
    GRDSTOCK.ColAlignment(6) = 1
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 4
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 4
    
    Txtdays.Text = 10
    PHY_FLAG = True
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    Left = 500
    Top = 0
    Call CMDDISPLAY_Click
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHY_FLAG = False Then PHY_REC.Close
   'Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_Click()
    CMBCHANGE.Visible = False
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 113, vbKeyReturn
            If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                Select Case GRDSTOCK.Col
                    Case 7
                        CMBCHANGE.Visible = True
                        CMBCHANGE.Top = GRDSTOCK.CellTop + 100
                        CMBCHANGE.Left = GRDSTOCK.CellLeft '+ 60
                        CMBCHANGE.Width = GRDSTOCK.CellWidth
                        'CmbPack.Height = GRDSTOCK.CellHeight
                        On Error Resume Next
                        CMBCHANGE.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                        CMBCHANGE.SetFocus
                End Select
            End If
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    CMBCHANGE.Visible = False
End Sub

Private Sub Txtdays_GotFocus()
    Txtdays.SelStart = 0
    Txtdays.SelLength = Len(Txtdays.Text)
End Sub

Private Sub Txtdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtdays.Text) = 0 Then Exit Sub
            Call CMDDISPLAY_Click
    End Select
End Sub

Private Sub Txtdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
