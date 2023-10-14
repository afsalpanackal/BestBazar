VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMROUTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ROUTE LAYOUT...."
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMROUTEWO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   18105
   Begin MSFlexGridLib.MSFlexGrid GRDARRANGE 
      Height          =   1020
      Left            =   300
      TabIndex        =   18
      Top             =   7575
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   1799
      _Version        =   393216
      Rows            =   1
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   8454143
      BackColorBkg    =   12632256
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
   Begin VB.Frame FRMEMAIN 
      Caption         =   "Frame1"
      Height          =   8685
      Left            =   15
      TabIndex        =   0
      Top             =   -270
      Width           =   18045
      Begin VB.CommandButton cmdRegister2 
         Caption         =   "Print Route Statement -2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   15390
         TabIndex        =   19
         Top             =   8055
         Width           =   1350
      End
      Begin VB.Frame frmesort 
         Height          =   6585
         Left            =   2895
         TabIndex        =   11
         Top             =   1350
         Visible         =   0   'False
         Width           =   5535
         Begin VB.CommandButton CMDDOWN 
            Caption         =   "Move Down"
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
            Left            =   1440
            TabIndex        =   17
            Top             =   6045
            Width           =   1245
         End
         Begin VB.CommandButton CMDUP 
            Caption         =   "Move UP"
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
            Left            =   105
            TabIndex        =   16
            Top             =   6045
            Width           =   1245
         End
         Begin VB.TextBox TXTsample 
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
            Height          =   290
            Left            =   2085
            TabIndex        =   15
            Top             =   3240
            Visible         =   0   'False
            Width           =   1005
         End
         Begin VB.CommandButton cmdcancel 
            Caption         =   "&Cancel"
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
            Left            =   4155
            TabIndex        =   14
            Top             =   6030
            Width           =   1230
         End
         Begin VB.CommandButton cmdOK 
            Caption         =   "&OK"
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
            Left            =   2850
            TabIndex        =   13
            Top             =   6030
            Width           =   1230
         End
         Begin MSFlexGridLib.MSFlexGrid grddummy 
            Height          =   5820
            Left            =   120
            TabIndex        =   12
            Top             =   150
            Width           =   5325
            _ExtentX        =   9393
            _ExtentY        =   10266
            _Version        =   393216
            Rows            =   1
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
      End
      Begin VB.CommandButton CMDSORT 
         Caption         =   "SORT"
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
         Left            =   75
         TabIndex        =   10
         Top             =   8070
         Width           =   1515
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
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
         Height          =   780
         Left            =   45
         TabIndex        =   5
         Top             =   165
         Width           =   17925
         Begin VB.OptionButton OptWS 
            BackColor       =   &H00C0C0FF&
            Caption         =   "WS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   8310
            TabIndex        =   25
            Top             =   285
            Width           =   1335
         End
         Begin VB.OptionButton OptRT 
            BackColor       =   &H00C0C0FF&
            Caption         =   "RT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   9690
            TabIndex        =   24
            Top             =   285
            Width           =   1350
         End
         Begin VB.OptionButton Optall 
            BackColor       =   &H00C0C0FF&
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   11130
            TabIndex        =   21
            Top             =   285
            Width           =   1425
         End
         Begin VB.OptionButton OptVan 
            BackColor       =   &H00C0C0FF&
            Caption         =   "VS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   330
            Left            =   6990
            TabIndex        =   20
            Top             =   285
            Value           =   -1  'True
            Width           =   1395
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1605
            TabIndex        =   6
            Top             =   240
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
            Left            =   3720
            TabIndex        =   7
            Top             =   255
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   688
            _Version        =   393216
            Format          =   123076609
            CurrentDate     =   40498
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "PERIOD FROM"
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
            Left            =   180
            TabIndex        =   9
            Top             =   330
            Width           =   1365
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
            Left            =   3315
            TabIndex        =   8
            Top             =   315
            Width           =   285
         End
      End
      Begin VB.CommandButton CMDREGISTER 
         Caption         =   "Print Route Statement -1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   13980
         TabIndex        =   4
         Top             =   8055
         Width           =   1350
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
         Height          =   525
         Left            =   16800
         TabIndex        =   2
         Top             =   8040
         Width           =   1155
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
         Height          =   525
         Left            =   12525
         TabIndex        =   1
         Top             =   8040
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6990
         Left            =   45
         TabIndex        =   3
         Top             =   960
         Width           =   17925
         _ExtentX        =   31618
         _ExtentY        =   12330
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   380
         BackColorFixed  =   0
         ForeColorFixed  =   8454143
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         SelectionMode   =   1
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   6
         Left            =   2925
         TabIndex        =   23
         Top             =   8100
         Width           =   2490
      End
      Begin VB.Label lblpvalue 
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
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   5565
         TabIndex        =   22
         Top             =   8040
         Width           =   3195
      End
   End
End
Attribute VB_Name = "FRMROUTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdcancel_Click()
    frmesort.Visible = False
    GRDTranx.SetFocus
End Sub

Private Sub CmDDisplay_Click()
    Dim RSTACTCODE As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim M As Long
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    lblpvalue.Caption = "0.00"
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass

    Set rstTRANX = New ADODB.Recordset
    If OptVan.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') AND BILL_TYPE ='V' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
        'rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SI' AND BILL_TYPE ='V' AND ACT_CODE <> '130000' ORDER BY AREA", DB, adOpenStatic, adLockReadOnly
    ElseIf OptWs.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') AND BILL_TYPE ='W' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    ElseIf OptRT.Value = True Then
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') AND BILL_TYPE ='R' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    End If
    
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = IIf(IsNull(rstTRANX!Area), "", rstTRANX!Area)
'        Set RSTACTCODE = New ADODB.Recordset
'        RSTACTCODE.Open "SELECT AREA FROM ACTMAST WHERE ACT_CODE = '" & rstTRANX!act_code & "'", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
'            GRDTranx.TextMatrix(M, 5) = RSTACTCODE!Area
'        End If
'        RSTACTCODE.Close
'        Set RSTACTCODE = Nothing
        GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf((IsNull(rstTRANX!BILL_ADDRESS) Or (rstTRANX!BILL_ADDRESS = "")), "", ", " & rstTRANX!BILL_ADDRESS)
        GRDTranx.TextMatrix(M, 6) = Format(Round(rstTRANX!VCH_AMOUNT - rstTRANX!DISCOUNT, 2), "0.00")
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    M = 0
    grddummy.rows = 1
    Set rstTRANX = New ADODB.Recordset
    If OptVan.Value = True Then
        rstTRANX.Open "SELECT DISTINCT Area From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') AND BILL_TYPE ='V' AND ACT_CODE <> '130000' ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT DISTINCT Area From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    End If
    Do Until rstTRANX.EOF
        M = M + 1
        grddummy.rows = grddummy.rows + 1
        grddummy.FixedRows = 1
        grddummy.TextMatrix(M, 0) = M
        grddummy.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!Area), "", rstTRANX!Area)
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTo.SetFocus
    End Select
End Sub

Private Sub CMDDOWN_Click()
    Dim GR, GC, i, n As Integer
    
    GR = grddummy.Row
    GC = grddummy.Col
    
    If grddummy.Row = grddummy.rows - 1 Then
        grddummy.Row = GR
        grddummy.Col = GC
        grddummy.SetFocus
        Exit Sub
    End If
    GRDARRANGE.rows = 1
    For i = 1 To grddummy.Row - 1
        GRDARRANGE.rows = GRDARRANGE.rows + 1
        GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(i, 0)
        GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(i, 1)
    Next
    GRDARRANGE.rows = GRDARRANGE.rows + 1
    GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(Val(grddummy.Row) + 1, 0)
    GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(Val(grddummy.Row) + 1, 1)
    
    i = i + 1
    GRDARRANGE.rows = GRDARRANGE.rows + 1
    GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(Val(grddummy.Row), 0)
    GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(Val(grddummy.Row), 1)
    
    n = i + 1
    For i = n To grddummy.rows - 1
        GRDARRANGE.rows = GRDARRANGE.rows + 1
        GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(i, 0)
        GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(i, 1)
    Next i
    
    Call ReArrangegrid
    grddummy.Row = GR + 1
    grddummy.Col = GC
    
     With grddummy
    .RowSel = .Row
    For i = 0 To 1
    .Col = i
    .CellBackColor = &H80C0FF
    Next i
    End With
    
    grddummy.SetFocus
End Sub

Private Sub cmdOK_Click()
    
    Dim rstTRANX As ADODB.Recordset
    Dim M, n As Long
    
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    For n = 1 To grddummy.rows - 1
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') AND AREA = '" & Trim(grddummy.TextMatrix(n, 1)) & "' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        
        Do Until rstTRANX.EOF
            M = M + 1
            GRDTranx.rows = GRDTranx.rows + 1
            GRDTranx.FixedRows = 1
            GRDTranx.TextMatrix(M, 0) = M
            GRDTranx.TextMatrix(M, 1) = rstTRANX!VCH_NO
            GRDTranx.TextMatrix(M, 2) = rstTRANX!VCH_DATE
            GRDTranx.TextMatrix(M, 3) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            GRDTranx.TextMatrix(M, 5) = rstTRANX!Area
            GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf((IsNull(rstTRANX!BILL_ADDRESS) Or (rstTRANX!BILL_ADDRESS = "")), "", ", " & rstTRANX!BILL_ADDRESS)
            GRDTranx.TextMatrix(M, 6) = Format(Round(rstTRANX!VCH_AMOUNT - rstTRANX!DISCOUNT, 2), "0.00")
            
            rstTRANX.MoveNext
        Loop
        rstTRANX.Close
        Set rstTRANX = Nothing
    Next n
    
    GRDTranx.Visible = True
    frmesort.Visible = False
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Sub

Private Sub CMDREGISTER_Click()
    
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTROUTE1"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OptVan.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='V')"
    ElseIf OptWs.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='W')"
    ElseIf OptRT.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='R')"
    Else
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO'))"
    End If
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
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
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
    
'    Call GENERATEREPORT
'    On Error GoTo CLOSEFILE
'    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
'CLOSEFILE:
'    If Err.Number = 55 Then
'        Close #1
'        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
'    End If
'    On Error GoTo eRRHAND
'
'    Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
'    Print #1, "EXIT"
'    Close #1
'
'    '//HERE write the proper path where your command.com file exist
'    'Shell "C:\WINDOW\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
'    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
'    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

'Private Function GENERATEREPORT()
'    Dim SN As Integer
'    Dim i As Long
'    SN = 0
'
'    On Error GoTo CLOSEFILE
'    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
'
'CLOSEFILE:
'    If Err.Number = 55 Then
'        Close #1
'        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
'    End If
'    On Error GoTo eRRHAND
'    '//Create Report Heading
'    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
'    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
'
'
'    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
'            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
'    'Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'
'    Print #1, Chr(27) & Chr(71) & Chr(10) & _
'              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
'    Print #1, Chr(27) & Chr(85) & Chr(0)
'    Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft(DTFROM.value & " To " & DTTO.value, 30) '& Chr(27) & Chr(72) & Space(2)
'
'    Print #1, Space(7) & RepeatString("-", 74)
'    '27,115,1 Fast printing
'    '27,85,0 Bi direction printing
'    'Chr(27) & Chr(72) double strike off
'
'    Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft("SL", 3) & Space(0) & _
'            AlignLeft("Route", 15) & Space(0) & _
'            AlignLeft("Bill#", 6) & Space(0) & _
'            AlignLeft("Customer", 20) & Space(0) & _
'            AlignRight("Amount", 11) & Space(0) & _
'            AlignLeft("|", 10) & Space(0) & _
'            AlignLeft("|", 10) & Space(0)
'
'    Print #1, Space(7) & RepeatString("-", 72)
'
'    For i = 1 To GRDTranx.Rows - 1
'        Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft(Str(i), 3) & Space(0) & _
'            AlignLeft(GRDTranx.TextMatrix(i, 5), 15) & _
'            AlignLeft(GRDTranx.TextMatrix(i, 1), 5) & _
'            AlignLeft(GRDTranx.TextMatrix(i, 3), 20) & _
'            AlignRight(Format(Round(GRDTranx.TextMatrix(i, 6), 10), "0.00"), 12) & Space(1) & _
'            AlignLeft("|", 10) & Space(0) & _
'            AlignLeft("|", 10) & Space(0) '& _
'            Chr(27) & Chr(72)  '//Bold Ends
'        Print #1, Space(7) & RepeatString("-", 74)
'    Next i
'
'    Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(48) & AlignRight(lblpvalue.Caption, 10)
'    Print #1,
'    Print #1, Space(7) & RepeatString("-", 74)
'    Print #1,
'    Print #1, Space(7) & RepeatString("-", 74)
'    Print #1,
'    Print #1, Space(7) & RepeatString("-", 74)
'    Print #1,
'    Print #1, Space(7) & RepeatString("-", 74)
'
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'    Print #1, Chr(13)
'
'    Close #1 '//Closing the file
'    Exit Function
'
'eRRHAND:
'    MsgBox Err.Description
'End Function

Private Function GENERATEREPORT2()
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
    Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft(DTFROM.Value & " To " & DTTo.Value, 30) '& Chr(27) & Chr(72) & Space(2)

    Print #1, Space(7) & RepeatString("-", 72)
    Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft("SL", 3) & Space(0) & _
            AlignLeft("Route", 15) & Space(0) & _
            AlignLeft("Customer", 24) & Space(0) & _
            Chr(27) & Chr(72)  '//Bold Ends

    Print #1, Space(7) & RepeatString("-", 72)
    
    For i = 1 To GRDTranx.rows - 1
        Print #1, Chr(27) & Chr(115) & Chr(1) & Chr(18) & Chr(27) & Chr(72) & Space(7) & AlignLeft(str(i), 3) & Space(0) & _
            AlignLeft(GRDTranx.TextMatrix(i, 5), 15) & _
            AlignLeft(GRDTranx.TextMatrix(i, 3), 20) & _
            Chr(27) & Chr(72)  '//Bold Ends
        Print #1, Space(7) & RepeatString("-", 72)
    Next i
    Print #1,
    Print #1, Space(7) & RepeatString("-", 72)
    Print #1,
    Print #1, Space(7) & RepeatString("-", 72)
    Print #1,
    Print #1, Space(7) & RepeatString("-", 72)
    Print #1,
    Print #1, Space(7) & RepeatString("-", 72)

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
    Print #1, Chr(13)

    Close #1 '//Closing the file
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function

Private Sub cmdRegister2_Click()
    Dim i As Long
    Screen.MousePointer = vbHourglass
    
    On Error GoTo ERRHAND
    ReportNameVar = Rptpath & "RPTROUTE2"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If OptVan.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='V')"
    ElseIf OptWs.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='W')"
    ElseIf OptRT.Value = True Then
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO') AND {TRXMAST.BILL_TYPE} ='R')"
    Else
        Report.RecordSelectionFormula = "({TRXMAST.VCH_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {TRXMAST.VCH_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND ({TRXMAST.TRX_TYPE}='GI' OR {TRXMAST.TRX_TYPE}='SI' or {TRXMAST.TRX_TYPE}='RI' or {TRXMAST.TRX_TYPE}='HI' OR {TRXMAST.TRX_TYPE}='WO'))"
    End If
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
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
    Next
    frmreport.Caption = "SALES REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CMDSORT_Click()
    If GRDTranx.rows <= 1 Then Exit Sub
    frmesort.Visible = True
    grddummy.SetFocus
End Sub

Private Sub CMDUP_Click()
    Dim GR, GC, i, n As Integer
    
    GR = grddummy.Row
    GC = grddummy.Col

    If grddummy.Row = 1 Or grddummy.rows <= 1 Then
        grddummy.Row = GR
        grddummy.Col = GC
        grddummy.SetFocus
        Exit Sub
    End If
    
    GRDARRANGE.rows = 1
    For i = 1 To grddummy.Row - 2
        GRDARRANGE.rows = GRDARRANGE.rows + 1
        GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(i, 0)
        GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(i, 1)
    Next i
    GRDARRANGE.rows = GRDARRANGE.rows + 1
    GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(Val(grddummy.Row), 0)
    GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(Val(grddummy.Row), 1)
    
    i = i + 1
    GRDARRANGE.rows = GRDARRANGE.rows + 1
    GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(Val(grddummy.Row) - 1, 0)
    GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(Val(grddummy.Row) - 1, 1)
    
    n = i + 1
    For i = n To grddummy.rows - 1
        GRDARRANGE.rows = GRDARRANGE.rows + 1
        GRDARRANGE.TextMatrix(i, 0) = grddummy.TextMatrix(i, 0)
        GRDARRANGE.TextMatrix(i, 1) = grddummy.TextMatrix(i, 1)
    Next i
    
    Call ReArrangegrid
    grddummy.Row = GR - 1
    grddummy.Col = GC
    grddummy.RowSel = grddummy.Row
    
    With grddummy
    .RowSel = .Row
    For i = 0 To 1
    .Col = i
    .CellBackColor = &H80C0FF
    Next i
    End With

    grddummy.SetFocus
End Sub

Private Sub DTFROM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            DTTo.SetFocus
    End Select
End Sub


Private Sub DTTO_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            DTFROM.SetFocus
    End Select
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    Call ReportGeneratION
    On Error GoTo CLOSEFILE
    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file
    End If
    On Error GoTo ERRHAND
    Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOW\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    Exit Sub
ERRHAND:
    MsgBox err.Description
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
    GRDTranx.TextMatrix(0, 3) = "CUSTOMER"
    GRDTranx.TextMatrix(0, 4) = "Bill Address"
    GRDTranx.TextMatrix(0, 5) = "Route"
    GRDTranx.TextMatrix(0, 6) = "BILL AMT"

    GRDTranx.ColWidth(0) = 750
    GRDTranx.ColWidth(1) = 1200
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 4350
    GRDTranx.ColWidth(4) = 5200
    GRDTranx.ColWidth(5) = 2300
    GRDTranx.ColWidth(6) = 2000
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 1
    GRDTranx.ColAlignment(5) = 1
    GRDTranx.ColAlignment(6) = 6
    
    grddummy.TextMatrix(0, 0) = "SL"
    grddummy.TextMatrix(0, 1) = "ROUTE"

    grddummy.ColWidth(0) = 750
    grddummy.ColWidth(1) = 4000
    
    grddummy.ColAlignment(0) = 3
    grddummy.ColAlignment(1) = 1
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 16125
    'Me.Height = 10125
    Me.Left = 0
    Me.Top = 0
    'txtPassword = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

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
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
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
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES SUMMARY FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT, 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Function ReportREGISTER()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim SN As Integer
    Dim TRXTOTAL As Double
    
    SN = 0
    TRXTOTAL = 0
    '//NOTE : Report file name should never contain blank space.
    db.Execute "delete From SALESREG2"
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
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
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) '& Chr(27) & Chr(72)
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!HO_NAME, 30)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "SALES REGSITER FOR THE PERIOD"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "FROM " & DTFROM.Value & " TO " & DTTo.Value
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    Set RSTTRXFILE = New ADODB.Recordset
    Print #1, Chr(27) & Chr(67) & Chr(0) & Space(13) & RepeatString("-", 59)
    Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
    Print #1, Space(12) & RepeatString("-", 59)
    SN = 0
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG2", db, adOpenStatic, adLockOptimistic, adCmdText
    'RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_NO", DB, adOpenStatic,adLockReadOnly
    RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='GI' OR TRX_TYPE='SI' or TRX_TYPE='RI' or TRX_TYPE='HI' OR TRX_TYPE='WO') ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        CMDDISPLAY.Tag = ""
        If RSTTRXFILE!SLSM_CODE = "A" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(Round((RSTTRXFILE!DISCOUNT * RSTTRXFILE!VCH_AMOUNT) / 100, 2), "0.00"))
        ElseIf RSTTRXFILE!SLSM_CODE = "P" Then
            CMDDISPLAY.Tag = IIf(IsNull(RSTTRXFILE!DISCOUNT), "", Format(RSTTRXFILE!DISCOUNT, "0.00"))
        End If
        CMDEXIT.Tag = ""
        CMDEXIT.Tag = IIf(IsNull(RSTTRXFILE!ADD_AMOUNT), "", RSTTRXFILE!ADD_AMOUNT)
        'SLIPAMT = SLIPAMT + RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(cmdview.Tag))
        Print #1, Chr(27) & Chr(71) & Space(5) & Chr(14) & Chr(15) & AlignRight(str(SN), 4) & ". " & Space(1) & _
            AlignLeft(RSTTRXFILE!VCH_DATE, 10) & _
            AlignRight(Format(Round(RSTTRXFILE!VCH_AMOUNT - (Val(CMDDISPLAY.Tag) + Val(CMDEXIT.Tag)), 0), "0.00"), 16)
        'Print #1, Chr(13)
        TRXTOTAL = TRXTOTAL + RSTTRXFILE!VCH_AMOUNT
        
        RSTSALEREG.AddNew
        RSTSALEREG!VCH_NO = RSTTRXFILE!VCH_NO
        RSTSALEREG!TRX_TYPE = "SI"
        RSTSALEREG!VCH_DATE = RSTTRXFILE!VCH_DATE
        RSTSALEREG!VCH_AMOUNT = RSTTRXFILE!VCH_AMOUNT
        RSTSALEREG!PAYAMOUNT = 0 ' TRXFILE!PAY_AMOUNT
        RSTSALEREG!ACT_NAME = "Sales"
        RSTSALEREG!ACT_CODE = "111001"
        RSTSALEREG!DISCOUNT = 0 'rstTRANX!DISCOUNT
        RSTSALEREG.Update
        
        RSTTRXFILE.MoveNext
    Loop
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1,
    
    Print #1, Chr(27) & Chr(71) & Chr(14) & Chr(15) & Space(13) & AlignRight("TOTAL AMOUNT", 12) & AlignRight((Format(TRXTOTAL, "####.00")), 11)
    Print #1, Space(56) & RepeatString("-", 16)
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Function

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
    
End Function

Private Function ReArrangegrid()
    Dim i As Long
    grddummy.rows = 1
    For i = 1 To GRDARRANGE.rows - 1
        grddummy.rows = grddummy.rows + 1
        grddummy.FixedRows = 1
        grddummy.TextMatrix(i, 0) = i
        grddummy.TextMatrix(i, 1) = GRDARRANGE.TextMatrix(i, 1)
    Next

End Function

Private Function CopyGrid()
    Dim i As Long
    grddummy.rows = 1
    For i = 1 To GRDTranx.rows - 1
        grddummy.rows = grddummy.rows + 1
        grddummy.TextMatrix(i, 0) = grddummy.TextMatrix(i, 0)
        grddummy.TextMatrix(i, 1) = grddummy.TextMatrix(i, 1)
        grddummy.TextMatrix(i, 2) = grddummy.TextMatrix(i, 2)
        grddummy.TextMatrix(i, 3) = grddummy.TextMatrix(i, 3)
        grddummy.TextMatrix(i, 4) = grddummy.TextMatrix(i, 4)
        grddummy.TextMatrix(i, 5) = grddummy.TextMatrix(i, 5)
        grddummy.TextMatrix(i, 6) = grddummy.TextMatrix(i, 6)
    Next
End Function

Private Sub grddummy_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 56
            Call CMDUP_Click
        Case 50
            Call CMDDOWN_Click
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grddummy.Col
                Case 0
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    grddummy.TextMatrix(grddummy.Row, grddummy.Col) = TXTsample.text
                    grddummy.Enabled = True
                    TXTsample.Visible = False
                    grddummy.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grddummy.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
       Case Asc("'"), Asc("["), Asc("]"), Asc("\")
           KeyAscii = 0
       Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
           KeyAscii = Asc(UCase(Chr(KeyAscii)))
       Case Else
           KeyAscii = 0
    End Select
End Sub

Private Sub GRDDUMMY_Click()
    'TXTsample.Visible = False
    'grddummy.SetFocus
End Sub

Private Sub GRDDUMMY_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim sitem As String
'    Dim i As lONG
    If grddummy.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
'            Select Case grddummy.Col
'                Case 0
'                    TXTsample.Visible = True
'                    TXTsample.Top = grddummy.CellTop + 150
'                    TXTsample.Left = grddummy.CellLeft + 135
'                    TXTsample.Width = grddummy.CellWidth - 25
'                    TXTsample.Text = grddummy.TextMatrix(grddummy.Row, grddummy.Col)
'                    TXTsample.SetFocus
'            End Select
            cmdOK_Click
        Case vbKeyEscape
            cmdcancel_Click
            'frmesort.Visible = False
            'GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDDUMMY_Scroll()
'    TXTsample.Visible = False
'    grddummy.SetFocus
End Sub


Private Function GENERATEREPORTtest()
    Dim SN As Integer
    Dim i As Long
    SN = 0
    
    On Error GoTo CLOSEFILE
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    End If
    On Error GoTo ERRHAND
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold


    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    'Print #1, Chr(13)
    'Print #1, Chr(13)
    
    Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(7) & Chr(14) & Chr(15) & Chr(27) & Chr(72)
    Print #1, Chr(27) & Chr(71) & Chr(14) & Space(7) & AlignLeft(DTFROM.Value & " To " & DTTo.Value, 30) '& Chr(27) & Chr(72) & Space(2)
    Print #1, Space(7) & RepeatString("-", 90)
    'Print #1, Chr(27) & Chr(72) & Chr(18) & Space(7) & AlignLeft("SL.", 3) & Space(0)
    Print #1, Chr(120) & Chr(0) & Chr(18) & Space(7) & AlignLeft("SL.", 3) & Space(0) & _
            AlignLeft("Route", 15) & Space(0) & _
            AlignLeft("Bill#", 6) & Space(0) & _
            AlignLeft("Customer", 24) & Space(0) & _
            AlignLeft("Amount", 9) & Space(0) & _
            AlignLeft("|", 18) & Space(0) & _
            AlignLeft("|", 18) & Space(0) '& _
            Chr(27) & Chr(72)  '//Bold Ends

    Print #1, Space(7) & RepeatString("-", 90)
    

    Close #1 '//Closing the file
    Exit Function

ERRHAND:
    MsgBox err.Description
End Function


