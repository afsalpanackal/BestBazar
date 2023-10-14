VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmexpStaffRegwo 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPENSE REGISTER"
   ClientHeight    =   10305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10305
   ScaleWidth      =   10830
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   615
      TabIndex        =   7
      Top             =   2475
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
      BackColor       =   &H0080FF80&
      Height          =   10560
      Left            =   -45
      TabIndex        =   0
      Top             =   -240
      Width           =   10905
      Begin VB.OptionButton OptEXP 
         BackColor       =   &H00C0FFFF&
         Caption         =   "EXPENSE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   2220
         TabIndex        =   24
         Top             =   255
         Width           =   1890
      End
      Begin VB.OptionButton OptMast 
         BackColor       =   &H00C0FFFF&
         Caption         =   "EMPLOYEE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   45
         TabIndex        =   23
         Top             =   255
         Value           =   -1  'True
         Width           =   2175
      End
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
         TabIndex        =   16
         Top             =   9060
         Visible         =   0   'False
         Width           =   1635
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "&EXIT"
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
         Left            =   6060
         TabIndex        =   5
         Top             =   9060
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
         Left            =   4725
         TabIndex        =   4
         Top             =   9075
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H0080FF80&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1185
         Left            =   105
         TabIndex        =   1
         Top             =   9060
         Width           =   8985
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
         Height          =   6360
         Left            =   60
         TabIndex        =   6
         Top             =   2685
         Width           =   10755
         _ExtentX        =   18971
         _ExtentY        =   11218
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
         BackColor       =   &H0080FF80&
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
         Top             =   660
         Width           =   10800
         Begin VB.Frame FrmEmployee 
            BackColor       =   &H0080C0FF&
            Caption         =   "EMPLOYEE"
            Height          =   1815
            Left            =   30
            TabIndex        =   20
            Top             =   180
            Width           =   5520
            Begin VB.OptionButton OPTPERIOD 
               BackColor       =   &H0080C0FF&
               Caption         =   "PERIOD FROM"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   75
               TabIndex        =   33
               Top             =   435
               Value           =   -1  'True
               Width           =   1605
            End
            Begin VB.OptionButton OPTCUSTOMER 
               BackColor       =   &H0080C0FF&
               Caption         =   "EMPLOYEE"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   90
               TabIndex        =   32
               Top             =   855
               Width           =   1320
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
               TabIndex        =   31
               Top             =   780
               Width           =   3720
            End
            Begin MSComCtl2.DTPicker DTFROM 
               Height          =   390
               Left            =   1725
               TabIndex        =   34
               Top             =   360
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               _Version        =   393216
               CalendarForeColor=   0
               CalendarTitleForeColor=   16576
               CalendarTrailingForeColor=   255
               Format          =   51052545
               CurrentDate     =   40498
            End
            Begin MSComCtl2.DTPicker DTTO 
               Height          =   390
               Left            =   3885
               TabIndex        =   35
               Top             =   375
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               _Version        =   393216
               Format          =   51052545
               CurrentDate     =   40498
            End
            Begin MSDataListLib.DataList Dlstemployee 
               Height          =   645
               Left            =   1710
               TabIndex        =   36
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
               TabIndex        =   21
               Top             =   435
               Width           =   285
            End
         End
         Begin VB.Frame Frmexpense 
            BackColor       =   &H008080FF&
            Caption         =   "EXPENSE"
            Height          =   1815
            Left            =   30
            TabIndex        =   19
            Top             =   180
            Visible         =   0   'False
            Width           =   5190
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
               Left            =   1335
               TabIndex        =   27
               Top             =   765
               Width           =   3765
            End
            Begin VB.OptionButton OPTEXPENSE 
               BackColor       =   &H008080FF&
               Caption         =   "EXPENSE"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   90
               TabIndex        =   26
               Top             =   840
               Width           =   1215
            End
            Begin VB.OptionButton Optempperiod 
               BackColor       =   &H008080FF&
               Caption         =   "PERIOD FROM"
               ForeColor       =   &H00000000&
               Height          =   210
               Left            =   105
               TabIndex        =   25
               Top             =   420
               Width           =   1605
            End
            Begin MSComCtl2.DTPicker DTFROMEMP 
               Height          =   390
               Left            =   1710
               TabIndex        =   28
               Top             =   330
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               _Version        =   393216
               CalendarForeColor=   0
               CalendarTitleForeColor=   16576
               CalendarTrailingForeColor=   255
               Format          =   51052545
               CurrentDate     =   40498
            End
            Begin MSComCtl2.DTPicker DTTOEMP 
               Height          =   390
               Left            =   3585
               TabIndex        =   29
               Top             =   330
               Width           =   1530
               _ExtentX        =   2699
               _ExtentY        =   688
               _Version        =   393216
               Format          =   51052545
               CurrentDate     =   40498
            End
            Begin MSDataListLib.DataList DataList2 
               Height          =   645
               Left            =   1335
               TabIndex        =   30
               Top             =   1125
               Width           =   3765
               _ExtentX        =   6641
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
               Index           =   4
               Left            =   3255
               TabIndex        =   22
               Top             =   390
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
Attribute VB_Name = "FrmexpStaffRegwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim EMP_REC As New ADODB.Recordset
Dim EMP_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub CMDDISPLAY_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    
    Dim FROMDATE As Date
    Dim TODATE As Date
    Dim I As Integer

    
    LBLTRXTOTAL.Caption = ""
    On Error GoTo eRRhAND
    
    FROMDATE = Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = Format(DTTO.Value, "MM,DD,YYYY")
    
    GRDTranx.Rows = 1
    If OPTCUSTOMER.Value = True And Dlstemployee.BoundText = "" Then
        MsgBox "Select Employee", vbOKOnly, "Expense Register"
        Exit Sub
    End If
    If OPTEXPENSE.Value = True And DataList2.BoundText = "" Then
        MsgBox "Select Expense Head", vbOKOnly, "Expense Register"
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    I = 0
    Set rstTRANX = New ADODB.Recordset
    If OptMast.Value = True Then
        If OPTPERIOD.Value = True Then
            rstTRANX.Open "SELECT * From TRXEXP_MAST WHERE [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # ORDER BY VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            GoTo MASTER
        ElseIf OPTCUSTOMER.Value = True Then
            rstTRANX.Open "SELECT * From TRXEXP_MAST WHERE [ACT_CODE] = '" & Dlstemployee.BoundText & "' ORDER BY VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            GoTo MASTER
        End If
    End If
    If OptEXP.Value = True Then
        If Optempperiod.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE_EXP WHERE [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # ORDER BY VCH_DATE,VCH_NO, LINE_NO", db2, adOpenStatic, adLockReadOnly
            GoTo EXPENSE
        ElseIf OPTEXPENSE.Value = True Then
            rstTRANX.Open "SELECT * From TRXFILE_EXP WHERE [EXP_CODE] = '" & DataList2.BoundText & "' ORDER BY VCH_DATE,VCH_NO, LINE_NO", db2, adOpenStatic, adLockReadOnly
            GoTo EXPENSE
        End If
    End If
    
MASTER:
    GRDTranx.TextMatrix(0, 3) = "Employee Name"
    GRDTranx.ColWidth(5) = 0
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        I = I + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(I, 0) = I
        GRDTranx.TextMatrix(I, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(I, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(I, 3) = rstTRANX!ACT_NAME
        GRDTranx.TextMatrix(I, 4) = Format(rstTRANX!VCH_AMOUNT, "0.00")
        GRDTranx.TextMatrix(I, 5) = ""
        GRDTranx.TextMatrix(I, 6) = "M"
        rstTRANX.MoveNext
    Loop
    GoTo SKIP
    
EXPENSE:
    GRDTranx.TextMatrix(0, 3) = "Expense Account"
    GRDTranx.ColWidth(5) = 1500
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        I = I + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(I, 0) = I
        GRDTranx.TextMatrix(I, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(I, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(I, 3) = rstTRANX!EXP_NAME
        GRDTranx.TextMatrix(I, 4) = Format(rstTRANX!TRX_TOTAL, "0.00")
        GRDTranx.TextMatrix(I, 5) = IIf(IsNull(rstTRANX!REMARKS), "", rstTRANX!REMARKS)
        GRDTranx.TextMatrix(I, 6) = "E"
        rstTRANX.MoveNext
    Loop
    GoTo SKIP
SKIP:
    GRDTranx.Visible = True
    If I > 22 Then GRDTranx.TopRow = I
    GRDTranx.SetFocus
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblemployee.Caption = ""
    empflag.Caption = ""
    Screen.MousePointer = vbDefault
    Exit Sub
    
eRRhAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
'    rptPRINT.ReportFileName = App.Path & "\RPTSALESREG.RPT"
'    rptPRINT.Formulas(0) = "PERIOD = '" & DTFROM.Value & " " & " TO " & " " & DTTO.Value & "'"
'    rptPRINT.Action = 1
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "Sl"
    GRDTranx.TextMatrix(0, 1) = "Voucher No"
    GRDTranx.TextMatrix(0, 2) = "Date"
    GRDTranx.TextMatrix(0, 3) = "Employee Name"
    GRDTranx.TextMatrix(0, 4) = "Amount"
    GRDTranx.TextMatrix(0, 5) = "Remarks"
    GRDTranx.TextMatrix(0, 6) = "Type"
    
    GRDTranx.ColWidth(0) = 1000
    GRDTranx.ColWidth(1) = 1400
    GRDTranx.ColWidth(2) = 1300
    GRDTranx.ColWidth(3) = 4500
    GRDTranx.ColWidth(4) = 1500
    GRDTranx.ColWidth(5) = 0
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
    
    CLOSEALL = 1
    ACT_FLAG = True
    EMP_FLAG = True
    Month (Date) - 2
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    Me.Width = 10845
    Me.Height = 11025
    Me.Left = 1500
    Me.Top = 0
    txtPassword = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
        If EMP_FLAG = False Then EMP_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
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

    Dim I As Integer
    Dim RSTTRXFILE As ADODB.Recordset

    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.Rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)

            GRDBILL.Rows = 1
            I = 0
            Set RSTTRXFILE = New ADODB.Recordset
            If GRDTranx.TextMatrix(GRDTranx.Row, 6) = "M" Then
                RSTTRXFILE.Open "Select * From TRXFILE_EXP WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " ORDER BY LINE_NO", db2, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTTRXFILE.EOF
                    I = I + 1
                    GRDBILL.Rows = GRDBILL.Rows + 1
                    GRDBILL.FixedRows = 1
                    GRDBILL.TextMatrix(I, 0) = I
                    GRDBILL.TextMatrix(I, 1) = RSTTRXFILE!EXP_NAME
                    GRDBILL.TextMatrix(I, 2) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                    GRDBILL.TextMatrix(I, 3) = IIf(IsNull(RSTTRXFILE!REMARKS), "", RSTTRXFILE!REMARKS)
                    RSTTRXFILE.MoveNext
                Loop
                GRDBILL.TextMatrix(0, 1) = "Expense Head"
                GRDBILL.ColWidth(0) = 1000
                GRDBILL.ColWidth(1) = 4500
                GRDBILL.ColWidth(2) = 1800
                GRDBILL.ColWidth(3) = 1800
            Else
                RSTTRXFILE.Open "Select * From TRXEXP_MAST WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & "", db2, adOpenStatic, adLockReadOnly, adCmdText
                Do Until RSTTRXFILE.EOF
                    I = I + 1
                    GRDBILL.Rows = GRDBILL.Rows + 1
                    GRDBILL.FixedRows = 1
                    GRDBILL.TextMatrix(I, 0) = I
                    GRDBILL.TextMatrix(I, 1) = RSTTRXFILE!ACT_NAME
                    GRDBILL.TextMatrix(I, 2) = Format(RSTTRXFILE!VCH_AMOUNT, "0.00")
                    RSTTRXFILE.MoveNext
                Loop
                GRDBILL.TextMatrix(0, 1) = "Employee Name"
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
    GRDTranx.Rows = 1
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub OptEXP_Click()
    FrmEmployee.Visible = False
    Frmexpense.Visible = True
    Optempperiod.Value = True
    OPTPERIOD.Value = False
    OPTCUSTOMER.Value = False
End Sub

Private Sub OPTEXPENSE_Click()
    'TXTDEALER.SetFocus
End Sub

Private Sub OptMast_Click()
    FrmEmployee.Visible = True
    Frmexpense.Visible = False
    OPTPERIOD.Value = True
    OPTEXPENSE.Value = False
    Optempperiod.Value = False
End Sub

Private Sub OPTPERIOD_GotFocus()
    LBLTRXTOTAL.Caption = ""
    GRDTranx.Rows = 1
End Sub

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='641')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='641')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    GRDTranx.Rows = 1
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub TXTEMPLOYEE_Change()
    On Error GoTo eRRhAND
    If empflag.Caption <> "1" Then
        If EMP_FLAG = True Then
            EMP_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='321')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTEMPLOYEE.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        Else
            EMP_REC.Close
            EMP_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='321')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTEMPLOYEE.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        End If
        If (EMP_REC.EOF And EMP_REC.BOF) Then
            lblemployee.Caption = ""
        Else
            lblemployee.Caption = EMP_REC!ACT_NAME
        End If
        Set Me.Dlstemployee.RowSource = EMP_REC
        Dlstemployee.ListField = "ACT_NAME"
        Dlstemployee.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTEMPLOYEE_GotFocus()
    TXTEMPLOYEE.SelStart = 0
    TXTEMPLOYEE.SelLength = Len(TXTEMPLOYEE.Text)
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
        Case Asc("'")
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
    Dlstemployee.Text = TXTEMPLOYEE.Text
    Call Dlstemployee_Click
End Sub

Private Sub Dlstemployee_LostFocus()
     empflag.Caption = ""
End Sub

Private Sub Dlstemployee_Click()
    TXTEMPLOYEE = Dlstemployee.Text
    lblemployee.Caption = TXTEMPLOYEE
End Sub

Private Sub Dlstemployee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Dlstemployee.Text = "" Then Exit Sub
            If IsNull(Dlstemployee.SelectedItem) Then
                MsgBox "Select Expense head From List", vbOKOnly, "Expense Entry..."
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

