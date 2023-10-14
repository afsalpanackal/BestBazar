VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMFREE 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ISSUE OF FREE ITEMS"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmfree.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   10755
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   9120
      Left            =   0
      TabIndex        =   0
      Top             =   -105
      Width           =   10875
      Begin VB.CommandButton CMDPRINTREGISTER 
         Caption         =   "PRINT REGISTER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   9060
         TabIndex        =   17
         Top             =   8580
         Width           =   1635
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7785
         TabIndex        =   5
         Top             =   8580
         Width           =   1200
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6495
         TabIndex        =   4
         Top             =   8580
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   465
         Left            =   105
         TabIndex        =   1
         Top             =   8565
         Width           =   8985
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL FREE QTY"
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
            Left            =   3150
            TabIndex        =   21
            Top             =   75
            Width           =   1590
         End
         Begin VB.Label LBLFREE 
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
            Height          =   450
            Left            =   4785
            TabIndex        =   20
            Top             =   0
            Width           =   1470
         End
         Begin VB.Label LBLSOLD 
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
            Height          =   450
            Left            =   1635
            TabIndex        =   3
            Top             =   15
            Width           =   1470
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL SOLD QTY"
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
            Left            =   0
            TabIndex        =   2
            Top             =   90
            Width           =   1590
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6105
         Left            =   75
         TabIndex        =   6
         Top             =   2445
         Width           =   10635
         _ExtentX        =   18759
         _ExtentY        =   10769
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2415
         Left            =   75
         TabIndex        =   7
         Top             =   15
         Width           =   10635
         Begin VB.CheckBox CHKFILTER 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFC0&
            Caption         =   "Add &Filter"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   4305
            MaskColor       =   &H00FFFFC0&
            TabIndex        =   31
            Top             =   2025
            Value           =   1  'Checked
            Width           =   1230
         End
         Begin VB.Frame Frmefilter 
            BackColor       =   &H00C0E0FF&
            Height          =   2250
            Left            =   5625
            TabIndex        =   22
            Top             =   135
            Width           =   4980
            Begin VB.OptionButton OptCompany 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Company"
               Height          =   210
               Left            =   45
               TabIndex        =   28
               Top             =   1260
               Value           =   -1  'True
               Width           =   1320
            End
            Begin VB.OptionButton OPTITEM 
               BackColor       =   &H00C0E0FF&
               Caption         =   "Item Name"
               Height          =   210
               Left            =   30
               TabIndex        =   27
               Top             =   195
               Width           =   1320
            End
            Begin VB.TextBox Txtcompany 
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
               Left            =   1425
               TabIndex        =   25
               Top             =   1215
               Width           =   3495
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
               Height          =   330
               Left            =   1410
               TabIndex        =   23
               Top             =   150
               Width           =   3495
            End
            Begin MSDataListLib.DataList DataList1 
               Height          =   645
               Left            =   1410
               TabIndex        =   24
               Top             =   495
               Width           =   3495
               _ExtentX        =   6165
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
            Begin MSDataListLib.DataList DataList3 
               Height          =   645
               Left            =   1425
               TabIndex        =   26
               Top             =   1560
               Width           =   3495
               _ExtentX        =   6165
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
         End
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
            Left            =   1845
            TabIndex        =   13
            Top             =   660
            Width           =   3720
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0FFC0&
            Caption         =   "CUSTOMER"
            Height          =   210
            Left            =   90
            TabIndex        =   12
            Top             =   705
            Width           =   1320
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0FFC0&
            Caption         =   "PERIOD FROM"
            Height          =   210
            Left            =   75
            TabIndex        =   11
            Top             =   255
            Value           =   -1  'True
            Width           =   1710
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   8
            Top             =   165
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   9
            Top             =   180
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1845
            TabIndex        =   14
            Top             =   1005
            Width           =   3720
            _ExtentX        =   6562
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
         Begin VB.Label lblcompanyflag 
            Height          =   315
            Left            =   480
            TabIndex        =   30
            Top             =   1965
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblcompany 
            Height          =   315
            Left            =   165
            TabIndex        =   29
            Top             =   1815
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label LBLMEDICINE 
            Height          =   315
            Left            =   105
            TabIndex        =   19
            Top             =   900
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label LBLMEDFLAG 
            Height          =   315
            Left            =   420
            TabIndex        =   18
            Top             =   1050
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   225
            TabIndex        =   16
            Top             =   1455
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   -90
            TabIndex        =   15
            Top             =   1305
            Visible         =   0   'False
            Width           =   1620
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
            Left            =   3585
            TabIndex        =   10
            Top             =   240
            Width           =   285
         End
      End
   End
End
Attribute VB_Name = "FRMFREE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim REPFLAG As Boolean 'REP
Dim RSTREP As New ADODB.Recordset
Dim REPCOMPflag As Boolean 'REP
Dim RSTCOMP As New ADODB.Recordset
Dim CLOSEALL As Integer
Dim selectionformla As String

Private Sub CHKFILTER_Click()
    If CHKFILTER.Value = 1 Then
        Frmefilter.Visible = 1
        tXTMEDICINE.SetFocus
    Else
        Frmefilter.Visible = 0
        tXTMEDICINE.text = ""
        txtcompany.text = ""
        If OPTPERIOD.Value = True Then
            DTFROM.SetFocus
        Else
            TXTDEALER.SetFocus
        End If
    End If
End Sub

Private Sub CmDDisplay_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim i As Long

    
    LBLSOLD.Caption = ""
    LBLFREE.Caption = ""
    On Error GoTo ERRHAND
    
    FROMDATE = Format(DTFROM.Value, "yyyy/mm/dd")
    TODATE = Format(DTTo.Value, "yyyy/mm/dd")
    
    GRDTranx.rows = 1
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then
        MsgBox "Select Customer from the List", vbOKOnly, "Report.."
        TXTDEALER.SetFocus
        Exit Sub
    End If
    If CHKFILTER.Value = 1 And OPTITEM.Value = True And DataList1.BoundText = "" Then
        MsgBox "Select Item from the List", vbOKOnly, "Report.."
        tXTMEDICINE.SetFocus
        Exit Sub
    End If
    
    If CHKFILTER.Value = 1 And OptCompany.Value = True And DataList3.BoundText = "" Then
        MsgBox "Select Company from the List", vbOKOnly, "Report.."
        txtcompany.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    i = 0
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        If CHKFILTER.Value = 1 And OPTITEM.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0 AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND ITEM_CODE = '" & DataList1.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # and {TRXFILE.ITEM_CODE}='" & DataList1.BoundText & "')"
            'Report.RecordSelectionFormula = "( {POSUB.QTY}-{POSUB.RCVD_QTY}<>0 and {POSUB.VCH_NO}= " & Val(txtBillNo.Text) & " )"
        ElseIf CHKFILTER.Value = 1 And OptCompany.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0 AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND MFGR = '" & DataList3.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # and {TRXFILE.MFGR}='" & DataList3.BoundText & "')"
        Else
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0 AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 AND {TRXFILE.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
        End If
    Else
        If CHKFILTER.Value = 1 And OPTITEM.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0  AND M_USER_ID = '" & DataList2.BoundText & "' AND ITEM_CODE = '" & DataList1.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 and {TRXFILE.M_USER_ID}= '" & DataList2.BoundText & "' and {TRXFILE.ITEM_CODE}='" & DataList1.BoundText & "')"
        ElseIf CHKFILTER.Value = 1 And OptCompany.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0  AND M_USER_ID = '" & DataList2.BoundText & "' AND MFGR = '" & DataList3.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 and {TRXFILE.M_USER_ID}= '" & DataList2.BoundText & "' and {TRXFILE.MFGR}='" & DataList3.BoundText & "')"
        Else
            rstTRANX.Open "SELECT * From TRXFILE WHERE FREE_QTY >0  AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
            selectionformla = "( {TRXFILE.FREE_QTY}>0 and {TRXFILE.M_USER_ID}= '" & DataList2.BoundText & "')"
        End If
    End If
    
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        i = i + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(i, 2) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "GI"
                GRDTranx.TextMatrix(i, 3) = "B2B"
            Case "HI"
                GRDTranx.TextMatrix(i, 3) = "B2C"
            Case "SV"
                GRDTranx.TextMatrix(i, 3) = "Service"
            Case "SI"
                GRDTranx.TextMatrix(i, 3) = "Wholesale"
            Case Else
                GRDTranx.TextMatrix(i, 3) = "Retail"
        End Select
        GRDTranx.TextMatrix(i, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(i, 5) = rstTRANX!ITEM_NAME
        GRDTranx.TextMatrix(i, 6) = rstTRANX!QTY
        LBLSOLD.Caption = Val(LBLSOLD.Caption) + Val(GRDTranx.TextMatrix(i, 6))
        GRDTranx.TextMatrix(i, 7) = rstTRANX!FREE_QTY
        LBLFREE.Caption = Val(LBLFREE.Caption) + Val(GRDTranx.TextMatrix(i, 7))
        GRDTranx.TextMatrix(i, 8) = Mid(rstTRANX!VCH_DESC, 15)
        rstTRANX.MoveNext
    Loop
    
    GRDTranx.Visible = True
    If i > 22 Then GRDTranx.TopRow = i
    GRDTranx.SetFocus
    rstTRANX.Close
    Set rstTRANX = Nothing
    

    flagchange.Caption = ""
    LBLMEDFLAG.Caption = ""
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    
    On Error GoTo ERRHAND
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    ReportNameVar = Rptpath & "Rptfreereg"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = selectionformla
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If OPTCUSTOMER.Value = True Then If CRXFormulaField.Name = "{@Customer}" Then CRXFormulaField.text = "'" & DataList2.text & "'"
        If CHKFILTER.Value = 1 And OPTITEM.Value = True Then If CRXFormulaField.Name = "{@Item}" Then CRXFormulaField.text = "'" & DataList1.text & "'"
        If CHKFILTER.Value = 1 And OptCompany.Value = True Then If CRXFormulaField.Name = "{@Item}" Then CRXFormulaField.text = "'" & DataList3.text & "'"
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "BILL NO"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "TYPE"
    GRDTranx.TextMatrix(0, 4) = "BILL DATE"
    GRDTranx.TextMatrix(0, 5) = "ITEM NAME"
    GRDTranx.TextMatrix(0, 6) = "SOLD QTY"
    GRDTranx.TextMatrix(0, 7) = "FREE QTY"
    GRDTranx.TextMatrix(0, 8) = "CUSTOMER"
    
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 0
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 2000
    GRDTranx.ColWidth(6) = 1100
    GRDTranx.ColWidth(7) = 1100
    GRDTranx.ColWidth(8) = 2500
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 1
    GRDTranx.ColAlignment(6) = 3
    GRDTranx.ColAlignment(7) = 1
    
    
    CLOSEALL = 1
    ACT_FLAG = True
    REPFLAG = True
    REPCOMPflag = True
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'CHKFILTER.Value = 1
    Me.Width = 10845
    Me.Height = 11025
    Me.Left = 1500
    Me.Top = 0
    txtPassword = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
        If REPFLAG = False Then RSTREP.Close
        If REPCOMPflag = False Then RSTCOMP.Close
        
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub OptCompany_Click()
    txtcompany.SetFocus
End Sub

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_GotFocus()
     LBLSOLD.Caption = ""
     LBLFREE.Caption = ""
    GRDTranx.rows = 1
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub


Private Sub OPTITEM_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub OPTPERIOD_GotFocus()
    LBLSOLD.Caption = ""
    LBLFREE.Caption = ""
    GRDTranx.rows = 1
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select

End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    LBLSOLD.Caption = ""
    LBLFREE.Caption = ""
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
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = "1"
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo ERRHAND
    If LBLMEDFLAG.Caption <> "1" Then
        If REPFLAG = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            REPFLAG = False
        Else
            RSTREP.Close
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            REPFLAG = False
        End If
        If (RSTREP.EOF And RSTREP.BOF) Then
            LBLMEDICINE.Caption = ""
        Else
            LBLMEDICINE.Caption = RSTREP!ITEM_NAME
        End If
        Set Me.DataList1.RowSource = RSTREP
        DataList1.ListField = "ITEM_NAME"
        DataList1.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
    OPTITEM.Value = True
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.VisibleCount = 0 Then Exit Sub
            DataList1.SetFocus
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

Private Sub DataList1_Click()
    tXTMEDICINE.text = DataList1.text
    GRDTranx.rows = 1
    LBLSOLD.Caption = ""
    LBLFREE.Caption = ""
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.text = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Item From List", vbOKOnly, "Report..."
                DataList1.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList1_GotFocus()
    LBLMEDFLAG.Caption = "1"
    tXTMEDICINE.text = LBLMEDICINE.Caption
    OPTITEM.Value = True
End Sub

Private Sub DataList1_LostFocus()
     LBLMEDFLAG.Caption = ""
End Sub

Private Sub txtcompany_Change()
    
    On Error GoTo ERRHAND
    If lblcompanyflag.Caption <> "1" Then
        If REPCOMPflag = True Then
            RSTCOMP.Open "Select DISTINCT MANUFACTURER From MANUFACT  WHERE MANUFACTURER Like '" & Me.txtcompany.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            REPCOMPflag = False
        Else
            RSTCOMP.Close
            RSTCOMP.Open "Select DISTINCT MANUFACTURER From MANUFACT  WHERE MANUFACTURER Like '" & Me.txtcompany.text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            REPCOMPflag = False
        End If
        If (RSTCOMP.EOF And RSTCOMP.BOF) Then
            lblcompany.Caption = ""
        Else
            lblcompany.Caption = RSTCOMP!MANUFACTURER
        End If
        Set Me.DataList3.RowSource = RSTCOMP
        DataList3.ListField = "MANUFACTURER"
        DataList3.BoundColumn = "MANUFACTURER"
    End If
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcompany_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.text)
    OptCompany.Value = True
End Sub

Private Sub txtcompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList3.VisibleCount = 0 Then Exit Sub
            DataList3.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList3_Click()
    txtcompany.text = DataList3.text
    GRDTranx.rows = 1
    LBLSOLD.Caption = ""
    LBLFREE.Caption = ""
End Sub

Private Sub DataList3_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList3.text = "" Then Exit Sub
            If IsNull(DataList3.SelectedItem) Then
                MsgBox "Select Company From List", vbOKOnly, "Report..."
                DataList3.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub DataList3_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList3_GotFocus()
    lblcompanyflag.Caption = "1"
    txtcompany.text = lblcompany.Caption
    OptCompany.Value = True
End Sub

Private Sub DataList3_LostFocus()
     lblcompanyflag.Caption = ""
End Sub

