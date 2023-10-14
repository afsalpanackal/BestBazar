VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmFAReg 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIXED ASSETS REGISTER"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMFAReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   12240
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   615
      TabIndex        =   7
      Top             =   1935
      Visible         =   0   'False
      Width           =   9780
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   45
         TabIndex        =   8
         Top             =   585
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColor       =   0
         ForeColor       =   16777215
         BackColorFixed  =   255
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
      Begin VB.Label LBLBILLAMT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   6930
         TabIndex        =   12
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT"
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
         Height          =   255
         Index           =   1
         Left            =   5940
         TabIndex        =   11
         Top             =   210
         Width           =   885
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "VOUCHER NO."
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
         Height          =   255
         Index           =   0
         Left            =   3405
         TabIndex        =   10
         Top             =   180
         Width           =   1275
      End
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4710
         TabIndex        =   9
         Top             =   180
         Width           =   1035
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00FFC0FF&
      Height          =   9060
      Left            =   -45
      TabIndex        =   0
      Top             =   -240
      Width           =   12285
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
         Left            =   9135
         TabIndex        =   16
         Top             =   1635
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
         Left            =   6870
         TabIndex        =   5
         Top             =   1635
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
         Left            =   5625
         TabIndex        =   4
         Top             =   1635
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   480
         Left            =   5595
         TabIndex        =   1
         Top             =   1020
         Width           =   4515
         Begin VB.Label LBLTRXTOTAL 
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
            Left            =   2205
            TabIndex        =   3
            Top             =   15
            Width           =   2220
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT"
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
            Left            =   195
            TabIndex        =   2
            Top             =   90
            Width           =   1935
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6855
         Left            =   60
         TabIndex        =   6
         Top             =   2145
         Width           =   12195
         _ExtentX        =   21511
         _ExtentY        =   12091
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00FFC0FF&
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
         Height          =   2025
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   12210
         Begin VB.Frame FrmEmployee 
            BackColor       =   &H0080C0FF&
            Caption         =   "FIXED ASSETS"
            Height          =   1815
            Left            =   30
            TabIndex        =   19
            Top             =   180
            Width           =   5520
            Begin VB.OptionButton OPTPERIOD 
               BackColor       =   &H0080C0FF&
               Caption         =   "PERIOD FROM"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   75
               TabIndex        =   23
               Top             =   435
               Value           =   -1  'True
               Width           =   1605
            End
            Begin VB.OptionButton OPTCUSTOMER 
               BackColor       =   &H0080C0FF&
               Caption         =   "Fixed Assets"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   90
               TabIndex        =   22
               Top             =   855
               Width           =   1545
            End
            Begin VB.TextBox TXTEMPLOYEE 
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
               Left            =   1710
               TabIndex        =   21
               Top             =   780
               Width           =   3720
            End
            Begin MSComCtl2.DTPicker DTFROM 
               Height          =   390
               Left            =   1725
               TabIndex        =   24
               Top             =   360
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
               Left            =   3885
               TabIndex        =   25
               Top             =   375
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               _Version        =   393216
               Format          =   123076609
               CurrentDate     =   40498
            End
            Begin MSDataListLib.DataList Dlstemployee 
               Height          =   645
               Left            =   1710
               TabIndex        =   26
               Top             =   1125
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
               ForeColor       =   &H00000000&
               Height          =   270
               Index           =   2
               Left            =   3405
               TabIndex        =   20
               Top             =   435
               Width           =   285
            End
         End
         Begin VB.Label lblemployee 
            Height          =   315
            Left            =   7200
            TabIndex        =   18
            Top             =   450
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label empflag 
            Height          =   315
            Left            =   7470
            TabIndex        =   17
            Top             =   150
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   225
            TabIndex        =   15
            Top             =   1455
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   -90
            TabIndex        =   14
            Top             =   1305
            Visible         =   0   'False
            Width           =   1620
         End
      End
   End
End
Attribute VB_Name = "FrmFAReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim EMP_REC As New ADODB.Recordset
Dim EMP_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim i As Long

    
    LBLTRXTOTAL.Caption = ""
    On Error GoTo ERRHAND
    
    FROMDATE = Format(DTFROM.Value, "yyyy/mm/dd")
    TODATE = Format(DTTo.Value, "yyyy/mm/dd")
    
    GRDTranx.rows = 1
    If OPTCUSTOMER.Value = True And Dlstemployee.BoundText = "" Then
        MsgBox "Select Assets", vbOKOnly, "Fixed Assets Register"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    i = 0
    Set rstTRANX = New ADODB.Recordset
'    If OptMast.Value = True Then
'        If OPTPERIOD.Value = True Then
'            rstTRANX.Open "SELECT * From TRXFXDASSETMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'            GoTo MASTER
'        ElseIf OPTCUSTOMER.Value = True Then
'            rstTRANX.Open "SELECT * From TRXFXDASSETMAST WHERE ACT_CODE = '" & Dlstemployee.BoundText & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'            GoTo MASTER
'        End If
'    End If
'    If OptEXP.Value = True Then
        If OPTPERIOD.Value = True Then
            rstTRANX.Open "SELECT * From TRXFXDASSETS WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From TRXFXDASSETS WHERE ACT_CODE = '" & Dlstemployee.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
'    End If
    'GRDTranx.ColWidth(5) = 0
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        i = i + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(i, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!VCH_AMOUNT, "0.00")
        GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!REMARKS), "", rstTRANX!REMARKS)
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        rstTRANX.MoveNext
    Loop

    GRDTranx.Visible = True
    If i > 22 Then GRDTranx.TopRow = i
    GRDTranx.SetFocus
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblemployee.Caption = ""
    empflag.Caption = ""
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    Dim i As Integer
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    ReportNameVar = Rptpath & "RPTFxdAssets"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OPTPERIOD.Value = True Then
        Report.RecordSelectionFormula = "({TRXFILE_EXP.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE_EXP.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    Else
        Report.RecordSelectionFormula = "({TRXFILE_EXP.ACT_CODE} = '" & Dlstemployee.BoundText & "' AND {TRXFILE_EXP.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXFILE_EXP.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # )"
    End If
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM TRXFXDASSETMAST ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM TRXFXDASSETS ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "Fixed Assets Register"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "Sl"
    GRDTranx.TextMatrix(0, 1) = "Vch No"
    GRDTranx.TextMatrix(0, 2) = "Date"
    GRDTranx.TextMatrix(0, 3) = "Expense Head"
    GRDTranx.TextMatrix(0, 4) = "Amount"
    GRDTranx.TextMatrix(0, 5) = "Remarks"
    GRDTranx.TextMatrix(0, 6) = "Type"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 1300
    GRDTranx.ColWidth(3) = 4000
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 3400
    GRDTranx.ColWidth(6) = 0
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 6
    GRDTranx.ColAlignment(5) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Expense Head"
    GRDBILL.TextMatrix(0, 2) = "Amount"
    GRDBILL.TextMatrix(0, 3) = "Remarks"
    
    GRDBILL.ColWidth(0) = 1000
    GRDBILL.ColWidth(1) = 4500
    GRDBILL.ColWidth(2) = 1800
    GRDBILL.ColWidth(3) = 1800
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(1) = 1
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 1
    
    ACT_FLAG = True
    EMP_FLAG = True
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 10845
    Me.Height = 11025
    Me.Left = 1500
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If EMP_FLAG = False Then EMP_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    FRMEMAIN.Enabled = True
    FRMEBILL.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    Exit Sub
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset

    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)

            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            If GRDTranx.TextMatrix(GRDTranx.Row, 6) = "M" Then
                RSTTRXFILE.Open "Select * From TRXFXDASSETS WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTTRXFILE.EOF
                    i = i + 1
                    GRDBILL.rows = GRDBILL.rows + 1
                    GRDBILL.FixedRows = 1
                    GRDBILL.TextMatrix(i, 0) = i
                    GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!EXP_NAME
                    GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                    GRDBILL.TextMatrix(i, 3) = IIf(IsNull(RSTTRXFILE!REMARKS), "", RSTTRXFILE!REMARKS)
                    RSTTRXFILE.MoveNext
                Loop
                GRDBILL.TextMatrix(0, 1) = "Expense Head"
                GRDBILL.ColWidth(0) = 1000
                GRDBILL.ColWidth(1) = 4500
                GRDBILL.ColWidth(2) = 1800
                GRDBILL.ColWidth(3) = 1800
            Else
                RSTTRXFILE.Open "Select * From TRXFXDASSETMAST WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "", db, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTTRXFILE.EOF
                    i = i + 1
                    GRDBILL.rows = GRDBILL.rows + 1
                    GRDBILL.FixedRows = 1
                    GRDBILL.TextMatrix(i, 0) = i
                    GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ACT_NAME
                    GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!VCH_AMOUNT, "0.00")
                    RSTTRXFILE.MoveNext
                Loop
                GRDBILL.TextMatrix(0, 1) = "Fixed Assets Head"
                GRDBILL.ColWidth(0) = 0
                GRDBILL.ColWidth(2) = 0
                GRDBILL.ColWidth(3) = 0
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEMAIN.Enabled = False
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

Private Sub OPTCUSTOMER_Click()
    'TXTEMPLOYEE.SetFocus
End Sub

Private Sub OPTCUSTOMER_GotFocus()
     LBLTRXTOTAL.Caption = ""
    GRDTranx.rows = 1
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub OPTPERIOD_GotFocus()
    LBLTRXTOTAL.Caption = ""
    GRDTranx.rows = 1
End Sub


Private Sub TXTEMPLOYEE_Change()
    On Error GoTo ERRHAND
    If empflag.Caption <> "1" Then
        If EMP_FLAG = True Then
            EMP_REC.Open "select ITEM_CODE, ITEM_NAME from ASTMAST  WHERE ITEM_NAME Like '" & Me.TXTEMPLOYEE.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        Else
            EMP_REC.Close
            EMP_REC.Open "select ITEM_CODE, ITEM_NAME from ASTMAST  WHERE ITEM_NAME Like '" & Me.TXTEMPLOYEE.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        End If
        If (EMP_REC.EOF And EMP_REC.BOF) Then
            lblemployee.Caption = ""
        Else
            lblemployee.Caption = EMP_REC!ITEM_NAME
        End If
        Set Me.Dlstemployee.RowSource = EMP_REC
        Dlstemployee.ListField = "ITEM_NAME"
        Dlstemployee.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTEMPLOYEE_GotFocus()
    TXTEMPLOYEE.SelStart = 0
    TXTEMPLOYEE.SelLength = Len(TXTEMPLOYEE.text)
End Sub

Private Sub TXTEMPLOYEE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Dlstemployee.VisibleCount = 0 Then Exit Sub
            Dlstemployee.SetFocus
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select

End Sub

Private Sub TXTEMPLOYEE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Dlstemployee_GotFocus()
    empflag.Caption = 1
    TXTEMPLOYEE = lblemployee.Caption
    Dlstemployee.text = TXTEMPLOYEE.text
    Call Dlstemployee_Click
End Sub

Private Sub Dlstemployee_LostFocus()
     empflag.Caption = ""
End Sub

Private Sub Dlstemployee_Click()
    OPTCUSTOMER.Value = True
    TXTEMPLOYEE = Dlstemployee.text
    lblemployee.Caption = TXTEMPLOYEE
End Sub

Private Sub Dlstemployee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Dlstemployee.text = "" Then Exit Sub
            If IsNull(Dlstemployee.SelectedItem) Then
                MsgBox "Select Expense head From List", vbOKOnly, "Assets Entry..."
                Dlstemployee.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTEMPLOYEE.SetFocus
    End Select
End Sub

Private Sub Dlstemployee_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

