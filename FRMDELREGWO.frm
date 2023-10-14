VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDELREGWO 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DELIVERY REGISTER"
   ClientHeight    =   10545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10755
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
   Icon            =   "FRMDELREGWO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10545
   ScaleWidth      =   10755
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   555
      TabIndex        =   7
      Top             =   1890
      Visible         =   0   'False
      Width           =   9780
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   9660
         _ExtentX        =   17039
         _ExtentY        =   7064
         _Version        =   393216
         Rows            =   1
         Cols            =   8
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
         Caption         =   "BILL AMT"
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
         Caption         =   "BILL NO."
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
         Left            =   3825
         TabIndex        =   10
         Top             =   180
         Width           =   840
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
      BackColor       =   &H00C0E0FF&
      Height          =   10560
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
         TabIndex        =   24
         Top             =   9255
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
         Top             =   9255
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
         Top             =   9255
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1185
         Left            =   105
         TabIndex        =   1
         Top             =   9255
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
         Height          =   7155
         Left            =   165
         TabIndex        =   6
         Top             =   2025
         Width           =   10680
         _ExtentX        =   18838
         _ExtentY        =   12621
         _Version        =   393216
         Rows            =   1
         Cols            =   10
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
         BackColor       =   &H00C0E0FF&
         Caption         =   "DELIVERYREGISTER"
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
         Height          =   1890
         Left            =   150
         TabIndex        =   13
         Top             =   90
         Width           =   10560
         Begin VB.Frame Frame1 
            BackColor       =   &H00C0FFC0&
            Height          =   840
            Left            =   7755
            TabIndex        =   25
            Top             =   960
            Width           =   2670
            Begin VB.OptionButton OPTALL 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Display All"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   30
               TabIndex        =   27
               Top             =   165
               Width           =   2190
            End
            Begin VB.OptionButton optpend 
               BackColor       =   &H00C0FFC0&
               Caption         =   "Display Pending"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   204
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   255
               Left            =   45
               TabIndex        =   26
               Top             =   525
               Value           =   -1  'True
               Width           =   2190
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
            TabIndex        =   20
            Top             =   825
            Width           =   3720
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0E0FF&
            Caption         =   "CUSTOMER"
            Height          =   210
            Left            =   90
            TabIndex        =   19
            Top             =   870
            Value           =   -1  'True
            Width           =   1320
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0E0FF&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   75
            TabIndex        =   18
            Top             =   420
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1860
            TabIndex        =   14
            Top             =   330
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   49086465
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   4035
            TabIndex        =   15
            Top             =   345
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   49086465
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1845
            TabIndex        =   21
            Top             =   1170
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
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   225
            TabIndex        =   23
            Top             =   1455
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   -90
            TabIndex        =   22
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
            TabIndex        =   17
            Top             =   405
            Width           =   285
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
            Left            =   1110
            TabIndex        =   16
            Top             =   405
            Width           =   555
         End
      End
   End
End
Attribute VB_Name = "FRMDELREGWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
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
    If OPTCUSTOMER.Value = True And DataList2.BoundText = "" Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    I = 0
    Select Case OPTALL.Value
        Case True
            If OPTPERIOD.Value = True Then
                Set rstTRANX = New ADODB.Recordset
                rstTRANX.Open "SELECT DISTINCT VCH_NO From TEMPCN WHERE [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            Else
                Set rstTRANX = New ADODB.Recordset
                rstTRANX.Open "SELECT DISTINCT VCH_NO From TEMPCN WHERE [ACT_CODE] = '" & DataList2.BoundText & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            End If
        Case Else
            If OPTPERIOD.Value = True Then
                Set rstTRANX = New ADODB.Recordset
                rstTRANX.Open "SELECT DISTINCT VCH_NO From TEMPCN WHERE [CHECK_FLAG] <>'Y' AND [VCH_DATE] <=# " & TODATE & " # AND [VCH_DATE] >=# " & FROMDATE & " # ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            Else
                Set rstTRANX = New ADODB.Recordset
                rstTRANX.Open "SELECT DISTINCT VCH_NO From TEMPCN WHERE [ACT_CODE] = '" & DataList2.BoundText & "' AND [CHECK_FLAG] <>'Y' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
            End If
    End Select
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        I = I + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(I, 0) = I
        GRDTranx.TextMatrix(I, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(I, 2) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "SI"
                GRDTranx.TextMatrix(I, 3) = "Wholesale"
            Case Else
                GRDTranx.TextMatrix(I, 3) = "Retail"
        End Select
        GRDTranx.TextMatrix(I, 4) = rstTRANX!VCH_DATE
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * From TEMPCN WHERE VCH_NO = " & rstTRANX!VCH_NO & "", db2, adOpenStatic, adLockReadOnly
        Do Until RSTTRXFILE.EOF
            GRDTranx.TextMatrix(I, 5) = Format(Val(GRDTranx.TextMatrix(I, 4)) + RSTTRXFILE!TRX_TOTAL, "0.00")
            GRDTranx.TextMatrix(I, 6) = IIf(IsNull(RSTTRXFILE!VCH_DESC), "", Mid(RSTTRXFILE!VCH_DESC, 15))
            GRDTranx.TextMatrix(I, 7) = IIf((RSTTRXFILE!CHECK_FLAG = "Y"), "", "Pending")
            GRDTranx.TextMatrix(I, 8) = IIf(IsNull(RSTTRXFILE!BILL_NO), "", RSTTRXFILE!BILL_NO)
            GRDTranx.TextMatrix(I, 9) = IIf(IsNull(RSTTRXFILE!BILL_DATE), "", RSTTRXFILE!BILL_DATE)
            LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + RSTTRXFILE!TRX_TOTAL, "0.00")
            RSTTRXFILE.MoveNext
        Loop
        GRDTranx.Col = 4
        GRDTranx.Row = I
        GRDTranx.CellForeColor = vbRed
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        rstTRANX.MoveNext
    Loop
    
    GRDTranx.Visible = True
    If I > 22 Then GRDTranx.TopRow = I
    GRDTranx.SetFocus
    rstTRANX.Close
    Set rstTRANX = Nothing
    

    flagchange.Caption = ""
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
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "DN NO"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "TYPE"
    GRDTranx.TextMatrix(0, 4) = "DN DATE"
    GRDTranx.TextMatrix(0, 5) = "AMOUNT"
    GRDTranx.TextMatrix(0, 6) = "DELIVERED TO"
    GRDTranx.TextMatrix(0, 7) = "STATUS"
    GRDTranx.TextMatrix(0, 8) = "INVOICE NO."
    GRDTranx.TextMatrix(0, 9) = "INVOICE DATE"
    
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 0
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1350
    GRDTranx.ColWidth(5) = 1200
    GRDTranx.ColWidth(6) = 2500
    GRDTranx.ColWidth(7) = 1100
    GRDTranx.ColWidth(8) = 1100
    GRDTranx.ColWidth(9) = 1100
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 6
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 1
    GRDTranx.ColAlignment(8) = 6
    GRDTranx.ColAlignment(9) = 6
    
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "MRP"
    GRDBILL.TextMatrix(0, 3) = "Rate"
    GRDBILL.TextMatrix(0, 4) = "Qty"
    GRDBILL.TextMatrix(0, 5) = "Amount"
    GRDBILL.TextMatrix(0, 6) = "Serial No"
    GRDBILL.TextMatrix(0, 7) = "Exp. Date"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2500
    GRDBILL.ColWidth(2) = 900
    GRDBILL.ColWidth(3) = 900
    GRDBILL.ColWidth(4) = 900
    GRDBILL.ColWidth(5) = 1100
    GRDBILL.ColWidth(6) = 1100
    GRDBILL.ColWidth(7) = 1200
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 6
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 6
    GRDBILL.ColAlignment(6) = 1
    
    CLOSEALL = 1
    ACT_FLAG = True
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
    
        If MDIMAIN.PCTMENU.Visible = True Then
            MDIMAIN.PCTMENU.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        Else
            MDIMAIN.pctmenu2.Enabled = True
            MDIMAIN.pctmenu2.SetFocus
        End If
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
            RSTTRXFILE.Open "Select * From TEMPCN WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 2)) & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                I = I + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(I, 0) = I
                GRDBILL.TextMatrix(I, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(I, 2) = Format(RSTTRXFILE!MRP, "0.00")
                GRDBILL.TextMatrix(I, 3) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(I, 4) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(I, 5) = Format(RSTTRXFILE!SALES_PRICE * RSTTRXFILE!QTY, "0.00")
                GRDBILL.TextMatrix(I, 6) = RSTTRXFILE!REF_NO
                GRDBILL.TextMatrix(I, 7) = IIf(IsNull(RSTTRXFILE!EXP_DATE), "", RSTTRXFILE!EXP_DATE)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            FRMEMAIN.Enabled = False
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

Private Sub OPTALL_GotFocus()
     LBLTRXTOTAL.Caption = ""
    GRDTranx.Rows = 1
End Sub

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
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

Private Sub optpend_GotFocus()
     LBLTRXTOTAL.Caption = ""
    GRDTranx.Rows = 1
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 2)='13')And (len(ACT_CODE)>2) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
            OPTPERIOD.SetFocus
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
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

