VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMPURCAHSEREGWO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE REPORT"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
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
   Icon            =   "FrmPurchaseregWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   11580
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   6360
      Left            =   1605
      TabIndex        =   12
      Top             =   1755
      Visible         =   0   'False
      Width           =   8160
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   5640
         Left            =   45
         TabIndex        =   13
         Top             =   645
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   9948
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
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   975
         TabIndex        =   23
         Top             =   270
         Width           =   3090
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
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
         Index           =   2
         Left            =   60
         TabIndex        =   22
         Top             =   285
         Width           =   900
      End
      Begin VB.Label LBLBILLAMT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   6915
         TabIndex        =   17
         Top             =   270
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
         Left            =   6000
         TabIndex        =   16
         Top             =   300
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
         Left            =   4140
         TabIndex        =   15
         Top             =   270
         Width           =   825
      End
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   4950
         TabIndex        =   14
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00FF80FF&
      Height          =   10095
      Left            =   -90
      TabIndex        =   0
      Top             =   -225
      Width           =   11670
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
         Height          =   450
         Left            =   8730
         TabIndex        =   11
         Top             =   9090
         Width           =   1530
      End
      Begin VB.CommandButton TMPDELETE 
         Caption         =   "TEMP. DELETE"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   10290
         TabIndex        =   10
         Top             =   9090
         Visible         =   0   'False
         Width           =   1320
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
         Height          =   450
         Left            =   7365
         TabIndex        =   9
         Top             =   9090
         Width           =   1290
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
         Height          =   450
         Left            =   6090
         TabIndex        =   8
         Top             =   9090
         Width           =   1200
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00FF80FF&
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
         Height          =   1770
         Left            =   150
         TabIndex        =   18
         Top             =   105
         Width           =   5550
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
            Left            =   1680
            TabIndex        =   5
            Top             =   720
            Width           =   3735
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00FF80FF&
            Caption         =   "CUSTOMER"
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   45
            TabIndex        =   4
            Top             =   720
            Width           =   1320
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00FF80FF&
            Caption         =   "PERIOD FROM"
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   45
            TabIndex        =   1
            Top             =   315
            Value           =   -1  'True
            Width           =   1620
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1680
            TabIndex        =   2
            Top             =   240
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            CalendarForeColor=   0
            CalendarTitleForeColor=   16576
            CalendarTrailingForeColor=   255
            Format          =   51838977
            CurrentDate     =   40498
         End
         Begin MSComCtl2.DTPicker DTTO 
            Height          =   390
            Left            =   3870
            TabIndex        =   3
            Top             =   240
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   51838977
            CurrentDate     =   40498
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1680
            TabIndex        =   6
            Top             =   1080
            Width           =   3735
            _ExtentX        =   6588
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
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   20
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   21
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
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
            Left            =   3405
            TabIndex        =   19
            Top             =   300
            Width           =   285
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6600
         Left            =   150
         TabIndex        =   7
         Top             =   1890
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   11642
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
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   330
         Left            =   5730
         TabIndex        =   24
         Tag             =   "5"
         Top             =   270
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         Picture         =   "FrmPurchaseregWO.frx":000C
         ForeColor       =   0
         BarPicture      =   "FrmPurchaseregWO.frx":0028
         Max             =   150
         Text            =   "PLEASE WAIT..."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Purchase"
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
         Height          =   480
         Index           =   6
         Left            =   180
         TabIndex        =   30
         Top             =   8565
         Width           =   1020
      End
      Begin VB.Label lblcrdt 
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
         Height          =   420
         Left            =   4560
         TabIndex        =   29
         Top             =   8595
         Width           =   2220
      End
      Begin VB.Label lblcash 
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
         Height          =   420
         Left            =   8040
         TabIndex        =   28
         Top             =   8550
         Width           =   2220
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Purcahse"
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
         Height          =   435
         Index           =   4
         Left            =   6945
         TabIndex        =   27
         Top             =   8520
         Width           =   1050
      End
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
         Height          =   420
         Left            =   1245
         TabIndex        =   26
         Top             =   8610
         Width           =   2220
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Purchase"
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
         Height          =   480
         Index           =   3
         Left            =   3420
         TabIndex        =   25
         Top             =   8565
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FRMPURCAHSEREGWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

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
            cmddisplay.SetFocus
    End Select
End Sub

Private Sub CMDDISPLAY_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TAXAMT, EXSALEAMT, TAXSALEAMT, MRPVALUE As Double
    Dim TAXRATE As Single
    
    db2.Execute "delete * From SALESREG"
    
    lblcrdt.Caption = "0.00"
    lblcash.Caption = "0.00"
    LBLTRXTOTAL.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.Rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        rstTRANX.Open "SELECT * From TRANSMASTWO WHERE [VCH_DATE] <=# " & DTTO.Value & " # AND [VCH_DATE] >=# " & DTFROM.Value & " # ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db2, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRANSMASTWO WHERE [ACT_CODE] = '" & DataList2.BoundText & "' AND [VCH_DATE] <=# " & DTTO.Value & " # AND [VCH_DATE] >=# " & DTFROM.Value & " # AND (TRX_TYPE='SI' OR TRX_TYPE='RI')", db2, adOpenStatic, adLockReadOnly
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db2, adOpenStatic, adLockOptimistic, adCmdText
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        rstTRANX.Open "SELECT * From TRANSMASTWO WHERE [VCH_DATE] <=# " & DTTO.Value & " # AND [VCH_DATE] >=# " & DTFROM.Value & " # AND TRX_TYPE='PI' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRANSMASTWO WHERE [ACT_CODE] = '" & DataList2.BoundText & "' AND [VCH_DATE] <=# " & DTTO.Value & " # AND [VCH_DATE] >=# " & DTFROM.Value & " # AND TRX_TYPE='PI' ORDER BY VCH_NO", db2, adOpenStatic, adLockReadOnly
    End If
    
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 3) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 5) = Format(Round(Val(GRDTranx.TextMatrix(M, 3)) - Val(GRDTranx.TextMatrix(M, 4)), 2), "0.00")
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        
        cmddisplay.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        If rstTRANX!TRX_TYPE <> "PI" Then GoTo SKIP
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select DISTINCT SALES_TAX From RTRXFILEWO WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            EXSALEAMT = 0
            TAXSALEAMT = 0
            TAXAMT = 0
            MRPVALUE = 0
            TAXRATE = RSTTRXFILE!SALES_TAX
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From RTRXFILEWO WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & "", db2, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                If RSTTRXFILE!SALES_TAX > 0 And RSTtax!CHECK_FLAG = "V" Then
                    TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    TAXAMT = TAXAMT + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                    
                Else
'                    If RSTtax!SALE_1_FLAG = "1" Then
'                        TAXAMT = TAXAMT + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
'                        MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
'                    End If
                    EXSALEAMT = EXSALEAMT + RSTtax!TRX_TOTAL
                End If
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            RSTSALEREG.AddNew
            RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
            RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
            RSTSALEREG!DISCOUNT = Val(GRDTranx.TextMatrix(M, 4))
            RSTSALEREG!VCH_AMOUNT = Val(GRDTranx.TextMatrix(M, 3))
            RSTSALEREG!PAYAMOUNT = 0
            RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
            Set RSTACTCODE = New ADODB.Recordset
            RSTACTCODE.Open "SELECT [KGST] FROM ACTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
                RSTSALEREG!TIN_NO = RSTACTCODE!KGST
            End If
            RSTACTCODE.Close
            Set RSTACTCODE = Nothing
            RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
            RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
            RSTSALEREG!TAXAMOUNT = TAXAMT
            RSTSALEREG!TAXRATE = TAXRATE
            cmddisplay.Tag = Val(cmddisplay.Tag) + EXSALEAMT
            FRMEMAIN.Tag = Val(FRMEMAIN.Tag) + TAXSALEAMT
            FRMEBILL.Tag = Val(FRMEBILL.Tag) + TAXAMT
            RSTSALEREG.Update
            
            RSTTRXFILE.MoveNext
        Loop
        
        GRDTranx.TextMatrix(M, 7) = Format(Val(cmddisplay.Tag), "0.00")
        GRDTranx.TextMatrix(M, 8) = Format(Val(FRMEMAIN.Tag), "0.00")
        GRDTranx.TextMatrix(M, 9) = Format(Val(FRMEBILL.Tag), "0.00")
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
         If (rstTRANX!POST_FLAG = "Y") Then
            lblcash.Caption = Format(Val(lblcash.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        Else
            lblcrdt.Caption = Format(Val(lblcrdt.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        End If
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    Screen.MousePointer = vbHourglass
    ReportNameVar = App.Path & "\RPTPURCHREG.rpt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.Count
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "G:\dbase\YEAR13-14\MEDINV.MDB", "admin", "!@#$%^&*())(*&^%$#@!"
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "PURCHASE REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "BILL NO"
    GRDTranx.TextMatrix(0, 2) = "BILL DATE"
    GRDTranx.TextMatrix(0, 3) = "BILL AMT"
    GRDTranx.TextMatrix(0, 4) = "DISC AMT"
    GRDTranx.TextMatrix(0, 5) = "NET AMT"
    GRDTranx.TextMatrix(0, 6) = "SUPPLIER"
    GRDTranx.TextMatrix(0, 7) = "EX. SALES"
    GRDTranx.TextMatrix(0, 8) = "TAX SALES"
    GRDTranx.TextMatrix(0, 9) = "TAX AMT"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 1200
    GRDTranx.ColWidth(2) = 1200
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 2000
    GRDTranx.ColWidth(7) = 1200
    GRDTranx.ColWidth(8) = 1200
    GRDTranx.ColWidth(9) = 1200
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 6
    GRDTranx.ColAlignment(4) = 6
    GRDTranx.ColAlignment(5) = 6
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 3
    GRDTranx.ColAlignment(8) = 3
    GRDTranx.ColAlignment(9) = 3
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    
    Month (Date) - 2
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.Value = Format(Date, "DD/MM/YYYY")
    Me.Width = 11670
    Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
    txtPassword = "YEAR " & Year(Date)
    ACT_FLAG = True
    CLOSEALL = 1
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

Private Sub txtPassword_GotFocus()
    txtPassword.Text = ""
    txtPassword.PasswordChar = " "
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
End Sub

Private Sub txtPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMBMONTH.SetFocus
    End Select
End Sub

Private Sub TXTPASSWORD_LostFocus()
    If UCase(txtPassword.Text) = "SARAKALAM" Then
        txtPassword = "YEAR " & Year(Date)
        txtPassword.PasswordChar = ""
        CMBMONTH.SetFocus
        Me.Height = 6000
    Else
        txtPassword = "YEAR " & Year(Date)
        txtPassword.PasswordChar = ""
        CMBMONTH.SetFocus
        Me.Height = 3700
    End If
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
    If FRMEBILL.Visible = True Then
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.Rows = 1 Then Exit Sub
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLSUPPLIER.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)
            GRDBILL.Rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From RTRXFILEWO WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'PI'", db2, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!PTR, "0.00")
                GRDBILL.TextMatrix(i, 3) = Format(RSTTRXFILE!P_DISC, "0.00")
                GRDBILL.TextMatrix(i, 4) = Format(RSTTRXFILE!SALES_TAX, "0.00")
                GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 6) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            Frmeperiod.Enabled = False
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    
    End Select
End Sub

Private Sub OPTCUSTOMER_Click()
    TXTDEALER.SetFocus
End Sub

Private Sub OPTCUSTOMER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            OPTPERIOD.SetFocus
    End Select
End Sub

Private Sub TMPDELETE_Click()
    If GRDTranx.Rows = 1 Then Exit Sub
    If MsgBox("Are You Sure You want to Delete BILL NO." & "*** " & GRDTranx.TextMatrix(GRDTranx.Row, 1) & " ****", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
    db2.Execute ("DELETE from [SALESREG] WHERE SALESREG.VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 1) & " AND SALESREG.TRX_TYPE = 'PI'")
    Call fillSTOCKREG
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
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly
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
eRRHAND:
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
            cmddisplay.SetFocus
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

Private Function fillSTOCKREG()
    Dim rstTRANX As ADODB.Recordset
    Dim TRX_AMOUNT As Double
    Dim i As Integer
    
    LBLTRXTOTAL.Caption = ""
    
    On Error GoTo eRRHAND
    TRX_AMOUNT = 0
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")

    Screen.MousePointer = vbHourglass
    TRX_AMOUNT = 0
    
    GRDTranx.Rows = 1
    i = 0
    count1 = 0
    
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    LBLTRXTOTAL.Caption = ""
    GRDTranx.Visible = False
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From SALESREG", db2, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        i = i + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 0) = i
        GRDTranx.TextMatrix(i, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(i, 2) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(i, 3) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!VCH_AMOUNT, "0.00")
        GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!DISCOUNT, "0.00")
        GRDTranx.TextMatrix(i, 6) = Format(rstTRANX!VCH_AMOUNT - rstTRANX!DISCOUNT, "0.00")
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        TRX_AMOUNT = TRX_AMOUNT + rstTRANX!VCH_AMOUNT
        rstTRANX.MoveNext
        vbalProgressBar1.Max = rstTRANX.RecordCount
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
    Loop
    
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")
    
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Function

Private Sub ReportGeneratION()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
   ' On Error GoTo errHand
    '//NOTE : Report file name should never contain blank space.
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)

    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!ADDRESS & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(52) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(52) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(52) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "Purchase Register for the Period"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "From " & DTFROM.Value & " TO " & DTTO.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 66)
    Print #1, Chr(27) & Chr(71) & Space(12) & AlignLeft("Sl", 3) & Space(2) & _
            AlignLeft("Supplier", 11) & Space(8) & _
            AlignLeft("Tin No", 11) & _
            AlignLeft("INV No", 8) & _
            AlignLeft("INV Date", 10) & Space(6) & _
            AlignLeft("INV Amt", 8) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 66)
    SN = 0
    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_DATE", db2, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        Print #1, Chr(27) & Chr(71) & Space(9) & AlignRight(Str(SN), 3) & "." & Space(1) & _
            AlignLeft(IIf(IsNull(RSTTRXFILE!ACT_NAME), "", Trim(RSTTRXFILE!ACT_NAME)), 18) & Space(1) & _
            AlignLeft(IIf(IsNull(RSTTRXFILE!TIN_NO), "", Trim(RSTTRXFILE!TIN_NO)), 10) & _
            AlignRight(IIf(IsNull(RSTTRXFILE!PINV), "", Trim(RSTTRXFILE!PINV)), 8) & Space(2) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(RSTTRXFILE!VCH_AMOUNT, "0.00"), 14)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(16) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(62) & RepeatString("-", 16)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
    'Print #1, Chr(27) & Chr(72) & Space(16) & AlignRight("**** WISH YOU A SPEEDY RECOVERY ****", 40)


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
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.txt" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTO.SetFocus
    End Select
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTO.SetFocus
    End Select
End Sub

Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmddisplay.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub
