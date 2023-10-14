VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmstkmvmreportvan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inward / Outward details of All Items (Branch Sale)"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15975
   Icon            =   "frmstkmvmreportvan.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   15975
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FRMEALLITEMS 
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   15945
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   3855
         TabIndex        =   11
         Top             =   45
         Width           =   5265
         Begin VB.OptionButton ptcategory 
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   14
            Top             =   450
            Value           =   -1  'True
            Width           =   1470
         End
         Begin VB.OptionButton OptCompany 
            Caption         =   "Company"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   45
            TabIndex        =   13
            Top             =   180
            Width           =   1470
         End
         Begin MSForms.ComboBox cmbcompany 
            Height          =   360
            Left            =   1530
            TabIndex        =   12
            Top             =   225
            Width           =   3690
            VariousPropertyBits=   746604571
            ForeColor       =   255
            DisplayStyle    =   7
            Size            =   "6509;635"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            BorderColor     =   255
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
      End
      Begin VB.OptionButton optWoCost 
         Caption         =   "Hide Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12015
         TabIndex        =   10
         Top             =   210
         Value           =   -1  'True
         Width           =   1515
      End
      Begin VB.OptionButton OptCost 
         Caption         =   "Show Value"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   12030
         TabIndex        =   9
         Top             =   495
         Width           =   1515
      End
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
         Height          =   480
         Left            =   14010
         TabIndex        =   8
         Top             =   210
         Width           =   1260
      End
      Begin VB.CommandButton CmdDisplay 
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
         Height          =   480
         Left            =   9450
         TabIndex        =   2
         Top             =   210
         Width           =   1110
      End
      Begin VB.CommandButton Cmdcancel 
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
         Height          =   480
         Left            =   10650
         TabIndex        =   1
         Top             =   210
         Width           =   1260
      End
      Begin MSFlexGridLib.MSFlexGrid GrdItems 
         Height          =   7545
         Left            =   30
         TabIndex        =   3
         Top             =   825
         Width           =   15870
         _ExtentX        =   27993
         _ExtentY        =   13309
         _Version        =   393216
         Rows            =   1
         Cols            =   11
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   8438015
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   3
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
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   315
         Left            =   675
         TabIndex        =   4
         Top             =   225
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   116391937
         CurrentDate     =   41640
         MinDate         =   40179
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   315
         Left            =   2280
         TabIndex        =   5
         Top             =   210
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   556
         _Version        =   393216
         Format          =   116391937
         CurrentDate     =   41640
         MinDate         =   40179
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Period"
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
         Height          =   270
         Index           =   7
         Left            =   60
         TabIndex        =   7
         Top             =   255
         Width           =   660
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
         Height          =   195
         Index           =   5
         Left            =   2010
         TabIndex        =   6
         Top             =   225
         Width           =   285
      End
   End
End
Attribute VB_Name = "frmstkmvmreportvan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmDDisplay_Click()
    If OptCost.Value = True Then
        GrdItems.TextMatrix(0, 4) = "OP Value"
        GrdItems.TextMatrix(0, 6) = "Inward Value"
        GrdItems.TextMatrix(0, 8) = "Outward Value"
        GrdItems.TextMatrix(0, 10) = "Value"
        GrdItems.ColWidth(4) = 1500
        GrdItems.ColWidth(6) = 1500
        GrdItems.ColWidth(8) = 1500
        GrdItems.ColWidth(10) = 1500
        
    Else
        GrdItems.TextMatrix(0, 4) = "" '"OP AMT"
        GrdItems.TextMatrix(0, 6) = "" '"INWARD AMT"
        GrdItems.TextMatrix(0, 8) = "" '"OUTWARD AMT"
        GrdItems.TextMatrix(0, 10) = ""
        GrdItems.ColWidth(4) = 0 '1500
        GrdItems.ColWidth(6) = 0 '1500
        GrdItems.ColWidth(8) = 0 '1500
        GrdItems.ColWidth(10) = 0 '1500
    End If
    
    Call Stkmvmnt_All_Items
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Function Stkmvmnt_All_Items()
    Dim OPQTY, OPVAL, RCVD_OP, RCVD_VAL As Double
    Dim INWARD, INWARD_VAL, OUTWARD, OUTWARD_VAL As Double
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim i As Long
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    i = 1
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    GrdItems.rows = 1
    Set RSTITEMMAST = New ADODB.Recordset
    If cmbcompany.ListIndex = -1 Then
        RSTITEMMAST.Open "SELECT *  FROM ITEMMASTVAN ORDER BY ITEM_NAME", db, adOpenForwardOnly
    Else
        If OptCompany.Value = True Then
            RSTITEMMAST.Open "SELECT * FROM ITEMMASTVAN WHERE MANUFACTURER = '" & cmbcompany.text & "'  ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTITEMMAST.Open "SELECT * FROM ITEMMASTVAN WHERE CATEGORY = '" & cmbcompany.text & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
    End If
    Do Until RSTITEMMAST.EOF
        OPQTY = 0
        OPVAL = 0
        RCVD_OP = 0
        RCVD_VAL = 0
        
        OPQTY = IIf(IsNull(RSTITEMMAST!OPEN_QTY), 0, RSTITEMMAST!OPEN_QTY)
        OPVAL = IIf(IsNull(RSTITEMMAST!OPEN_VAL), 0, RSTITEMMAST!OPEN_VAL)
        
        INWARD = 0
        OUTWARD = 0
        INWARD_VAL = 0
        OUTWARD_VAL = 0
        
        'OPENING QTY
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(QTY + FREE_QTY) FROM RTRXFILEVAN where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            OPQTY = OPQTY + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing

        If OptCost.Value = True Then
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) FROM RTRXFILEVAN where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                OPVAL = OPVAL + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
        End If
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILEVAN WHERE ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            RCVD_OP = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            'RCVD_OP = RCVD_OP + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            'OUTWARD_VAL = OUTWARD_VAL + IIf(IsNull(rststock!TRX_TOTAL), 0, rststock!TRX_TOTAL)
            'rststock.MoveNext
        End If
        rststock.Close
        Set rststock = Nothing
        
        If OptCost.Value = True Then
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM((QTY + FREE_QTY) * ITEM_COST) FROM TRXFILEVAN WHERE ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE <'" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                RCVD_VAL = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
        End If
        
        OPQTY = OPQTY - RCVD_OP
        OPVAL = OPVAL - RCVD_VAL
        
        'INWARD QTY
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(QTY + FREE_QTY) FROM RTRXFILEVAN where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
        If OptCost.Value = True Then
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM(TRX_TOTAL) FROM RTRXFILEVAN where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                INWARD_VAL = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
        End If
        
        
        'OUTWARD
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILEVAN WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            OUTWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            'OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
            'OUTWARD_VAL = OUTWARD_VAL + IIf(IsNull(rststock!TRX_TOTAL), 0, rststock!TRX_TOTAL)
            'rststock.MoveNext
        End If
        rststock.Close
        Set rststock = Nothing
        
        If OptCost.Value = True Then
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT SUM((QTY + FREE_QTY) * ITEM_COST) FROM TRXFILEVAN WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "'", db, adOpenForwardOnly
            If Not (rststock.EOF And rststock.BOF) Then
                OUTWARD_VAL = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            End If
            rststock.Close
            Set rststock = Nothing
        End If
        
        GrdItems.rows = GrdItems.rows + 1
        GrdItems.FixedRows = 1
        GrdItems.TextMatrix(i, 0) = i
        GrdItems.TextMatrix(i, 1) = RSTITEMMAST!ITEM_CODE
        GrdItems.TextMatrix(i, 2) = RSTITEMMAST!ITEM_NAME
        GrdItems.TextMatrix(i, 3) = OPQTY
        GrdItems.TextMatrix(i, 4) = Round(OPVAL, 2)
        GrdItems.TextMatrix(i, 5) = INWARD
        GrdItems.TextMatrix(i, 6) = Round(INWARD_VAL, 2)
        GrdItems.TextMatrix(i, 7) = OUTWARD
        GrdItems.TextMatrix(i, 8) = Round(OUTWARD_VAL, 2)
        GrdItems.TextMatrix(i, 9) = Round(OPQTY + INWARD - OUTWARD, 2)
        GrdItems.TextMatrix(i, 10) = Round(OPVAL + INWARD_VAL - OUTWARD_VAL, 2)
        i = i + 1
        
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
        
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub CmdExport_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then MsgBox "Permission Denied", vbOKOnly, "STOCK MOVEMENT"
    If MsgBox("Are you sure?", vbYesNo, "STOCK MOVEMENT") = vbNo Then Exit Sub
    Dim oApp As Excel.Application
    Dim oWB As Excel.Workbook
    Dim oWS As Excel.Worksheet
    Dim xlRange As Excel.Range
    Dim i, n As Long
    
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
        oWS.Range("A1", "K1").Merge
        oWS.Range("A1", "K1").HorizontalAlignment = xlCenter
        oWS.Range("A2", "K2").Merge
        oWS.Range("A2", "K2").HorizontalAlignment = xlCenter
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
    oWS.Range("A" & 2).Value = "STOCK MOVEMENT REPORT"
    
    
    'oApp.Selection.Font.Bold = False
    oWS.Range("A" & 3).Value = GrdItems.TextMatrix(0, 0)
    oWS.Range("B" & 3).Value = GrdItems.TextMatrix(0, 1)
    oWS.Range("C" & 3).Value = GrdItems.TextMatrix(0, 2)
    oWS.Range("D" & 3).Value = GrdItems.TextMatrix(0, 3)
    On Error Resume Next
    oWS.Range("E" & 3).Value = GrdItems.TextMatrix(0, 4)
    oWS.Range("F" & 3).Value = GrdItems.TextMatrix(0, 5)
    oWS.Range("G" & 3).Value = GrdItems.TextMatrix(0, 6)
    oWS.Range("H" & 3).Value = GrdItems.TextMatrix(0, 7)
    oWS.Range("I" & 3).Value = GrdItems.TextMatrix(0, 8)
    oWS.Range("J" & 3).Value = GrdItems.TextMatrix(0, 9)
    oWS.Range("K" & 3).Value = GrdItems.TextMatrix(0, 10)
    On Error GoTo ErrHand
    
    i = 4
    For n = 1 To GrdItems.rows - 1
        oWS.Range("A" & i).Value = GrdItems.TextMatrix(n, 0)
        oWS.Range("B" & i).Value = GrdItems.TextMatrix(n, 1)
        oWS.Range("C" & i).Value = GrdItems.TextMatrix(n, 2)
        oWS.Range("D" & i).Value = GrdItems.TextMatrix(n, 3)
        oWS.Range("E" & i).Value = GrdItems.TextMatrix(n, 4)
        oWS.Range("F" & i).Value = GrdItems.TextMatrix(n, 5)
        oWS.Range("G" & i).Value = GrdItems.TextMatrix(n, 6)
        oWS.Range("H" & i).Value = GrdItems.TextMatrix(n, 7)
        oWS.Range("I" & i).Value = GrdItems.TextMatrix(n, 8)
        oWS.Range("J" & i).Value = GrdItems.TextMatrix(n, 9)
        oWS.Range("K" & i).Value = GrdItems.TextMatrix(n, 10)
        On Error GoTo ErrHand
        i = i + 1
    Next n
    oWS.Range("A" & i, "Z" & i).Select                      '-- particular cell selection
    oWS.Columns("A:Z").EntireColumn.AutoFit
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

Private Sub Form_Load()
    
     
    OptCompany.Value = True
    Call OptCompany_Click
    GrdItems.TextMatrix(0, 0) = "SL"
    GrdItems.TextMatrix(0, 1) = "ITEM CODE"
    GrdItems.TextMatrix(0, 2) = "ITEM NAME"
    GrdItems.TextMatrix(0, 3) = "OP QTY"
    GrdItems.TextMatrix(0, 4) = "" '"OP AMT"
    GrdItems.TextMatrix(0, 5) = "INWARD QTY"
    GrdItems.TextMatrix(0, 6) = "" '"INWARD AMT"
    GrdItems.TextMatrix(0, 7) = "OUTWARD QTY"
    GrdItems.TextMatrix(0, 8) = "" '"OUTWARD AMT"
    GrdItems.TextMatrix(0, 9) = "BALANCE"
   
    
    GrdItems.ColWidth(0) = 400
    GrdItems.ColWidth(1) = 1600
    GrdItems.ColWidth(2) = 3500
    GrdItems.ColWidth(3) = 1100
    GrdItems.ColWidth(4) = 0 '1500
    GrdItems.ColWidth(5) = 1100
    GrdItems.ColWidth(6) = 0 '1500
    GrdItems.ColWidth(7) = 1100
    GrdItems.ColWidth(8) = 0 '1500
    GrdItems.ColWidth(9) = 1100
    
    
    GrdItems.ColAlignment(0) = 1
    GrdItems.ColAlignment(1) = 1
    GrdItems.ColAlignment(2) = 1
    GrdItems.ColAlignment(3) = 1
    GrdItems.ColAlignment(4) = 1
    GrdItems.ColAlignment(5) = 4
    GrdItems.ColAlignment(6) = 4
    GrdItems.ColAlignment(7) = 1
     GrdItems.ColAlignment(8) = 1
    GrdItems.ColAlignment(9) = 4
    
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
    'Me.Height = 9990
    'Me.Width = 18555
    Me.Left = 0
    Me.Top = 0
    
End Sub

Private Sub OptCompany_Click()
    cmbcompany.Clear
    On Error GoTo ErrHand
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT DISTINCT MANUFACTURER FROM ITEMMAST  ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!MANUFACTURER) Then cmbcompany.AddItem (RSTCOMPANY!MANUFACTURER)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub ptcategory_Click()
    cmbcompany.Clear
    On Error GoTo ErrHand
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT DISTINCT CATEGORY FROM ITEMMAST  ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!Category) Then cmbcompany.AddItem (RSTCOMPANY!Category)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub
