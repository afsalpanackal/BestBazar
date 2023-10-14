VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMAREARPT 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "AREA WISE STOCK WISE REPORT"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   ControlBox      =   0   'False
   Icon            =   "FRMAREAWISEWO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   9855
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   105
      TabIndex        =   8
      Top             =   -135
      Width           =   9705
      Begin VB.OptionButton OptCust 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Customer"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4500
         TabIndex        =   18
         Top             =   885
         Width           =   1200
      End
      Begin VB.OptionButton OPTAREA 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   360
         Left            =   4500
         TabIndex        =   17
         Top             =   405
         Value           =   -1  'True
         Width           =   780
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
         Left            =   90
         TabIndex        =   10
         Top             =   390
         Width           =   4260
      End
      Begin VB.CheckBox chkfrom 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Period From"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   4500
         TabIndex        =   9
         Top             =   1365
         Width           =   1335
      End
      Begin MSDataListLib.DataList DataList1 
         Height          =   1035
         Left            =   105
         TabIndex        =   11
         Top             =   750
         Width           =   4245
         _ExtentX        =   7488
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
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   5895
         TabIndex        =   12
         Top             =   1335
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   110690305
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   7905
         TabIndex        =   13
         Top             =   1335
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   110690305
         CurrentDate     =   40498
      End
      Begin MSForms.ComboBox cmbcustomer 
         Height          =   360
         Left            =   5715
         TabIndex        =   19
         Top             =   870
         Width           =   3810
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "6720;635"
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
      Begin MSForms.ComboBox cmbarea 
         Height          =   360
         Left            =   5700
         TabIndex        =   16
         Top             =   390
         Width           =   3810
         VariousPropertyBits=   746604571
         ForeColor       =   255
         DisplayStyle    =   7
         Size            =   "6720;635"
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
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
         Index           =   2
         Left            =   105
         TabIndex        =   15
         Top             =   165
         Width           =   4200
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   5
         Left            =   7575
         TabIndex        =   14
         Top             =   1410
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdshow 
      Caption         =   "&Show"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6135
      TabIndex        =   7
      Top             =   8775
      Width           =   1125
   End
   Begin VB.CommandButton cmdprint 
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
      Height          =   465
      Left            =   7410
      TabIndex        =   3
      Top             =   8775
      Width           =   1125
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
      Height          =   480
      Left            =   8670
      TabIndex        =   0
      Top             =   8760
      Width           =   1125
   End
   Begin MSFlexGridLib.MSFlexGrid grdsTOCK 
      Height          =   6915
      Left            =   120
      TabIndex        =   6
      Top             =   1800
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   12197
      _Version        =   393216
      Rows            =   1
      Cols            =   9
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      GridLineWidth   =   2
   End
   Begin VB.Label lbldealer 
      BackColor       =   &H00FF80FF&
      Height          =   315
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      BackColor       =   &H00FF80FF&
      Height          =   315
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLSTAOCKVALUE 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   2205
      TabIndex        =   2
      Top             =   8730
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Sold Qty"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   300
      Index           =   1
      Left            =   150
      TabIndex        =   1
      Top             =   8790
      Width           =   2235
   End
End
Attribute VB_Name = "FRMAREARPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSTREP As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim REPFLAG As Boolean 'RE

Private Sub cmbarea_GotFocus()
    OPTAREA.value = True
End Sub

Private Sub cmbcustomer_Click()
    OptCust.value = True
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdshow_Click()
    Dim rststock As ADODB.Recordset

    Dim i As Long
    Dim N As Double
    Dim FROMDATE As Date
    Dim TODATE As Date
    
    If OPTAREA.value = True And cmbarea.ListIndex = -1 Then
        MsgBox "Select Area from the list", vbOKOnly, "STOCK MOVEMENT"
        Exit Sub
    End If
    If OptCust.value = True And cmbcustomer.ListIndex = -1 Then
        MsgBox "Select Customer from the list", vbOKOnly, "STOCK MOVEMENT"
        Exit Sub
    End If
    FROMDATE = DTFROM.value 'Format(DTFROM.Value, "MM,DD,YYYY")
    TODATE = DTTO.value 'Format(DTTO.Value, "MM,DD,YYYY")
    
    i = 0
    N = 0
    Screen.MousePointer = vbHourglass
    GRDSTOCK.Rows = 1
    On Error GoTo Errhand
    GRDSTOCK.ColWidth(0) = 300
    GRDSTOCK.ColWidth(1) = 600
    GRDSTOCK.ColWidth(2) = 2650
    GRDSTOCK.ColWidth(3) = 500
    GRDSTOCK.ColWidth(4) = 500
    GRDSTOCK.ColWidth(5) = 800
    GRDSTOCK.ColWidth(6) = 800
    GRDSTOCK.ColWidth(7) = 800
    GRDSTOCK.ColWidth(8) = 1800
    
    GRDSTOCK.ColAlignment(0) = 4
    'grdsTOCK.ColAlignment(1) = 4
    'grdsTOCK.ColAlignment(2) = 4
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 7
    GRDSTOCK.ColAlignment(5) = 7
    GRDSTOCK.ColAlignment(6) = 7
    GRDSTOCK.ColAlignment(7) = 7
    'grdsTOCK.ColAlignment(8) = 7
    
    GRDSTOCK.TextArray(0) = "SL"
    GRDSTOCK.TextArray(1) = "ITEM CODE"
    GRDSTOCK.TextArray(2) = "ITEM NAME"
    GRDSTOCK.TextArray(3) = "UNIT"
    GRDSTOCK.TextArray(4) = "QTY"
    GRDSTOCK.TextArray(5) = "FREE"
    GRDSTOCK.TextArray(6) = "MRP"
    GRDSTOCK.TextArray(7) = "RATE"
    GRDSTOCK.TextArray(8) = "CUSTOMER"
    
    Set rststock = New ADODB.Recordset
    If chkfrom.value = 1 Then
        If OPTAREA.value = True Then
            rststock.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRXFILE.ITEM_CODE = '" & DataList1.BoundText & "'AND TRXFILE.AREA = '" & cmbarea.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            rststock.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRXFILE.ITEM_CODE = '" & DataList1.BoundText & "'AND MID(TRXFILE.VCH_DESC,15) = '" & cmbcustomer.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If OPTAREA.value = True Then
            rststock.Open "Select * From TRXFILE WHERE TRXFILE.ITEM_CODE = '" & DataList1.BoundText & "'AND TRXFILE.AREA = '" & cmbarea.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        Else
            rststock.Open "Select * From TRXFILE WHERE TRXFILE.ITEM_CODE = '" & DataList1.BoundText & "'AND MID(TRXFILE.VCH_DESC,15) = '" & cmbcustomer.Text & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        End If
    End If
    
    Do Until rststock.EOF
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        i = i + 1
                
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = rststock!UNIT
        GRDSTOCK.TextMatrix(i, 4) = Val(rststock!QTY)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!FREE_QTY), 0, Val(rststock!FREE_QTY))
        GRDSTOCK.TextMatrix(i, 6) = Format(rststock!MRP, ".000")
        GRDSTOCK.TextMatrix(i, 7) = Format(rststock!SALES_PRICE, ".000")
        GRDSTOCK.TextMatrix(i, 8) = Mid(rststock!VCH_DESC, 15)
        N = N + Val(GRDSTOCK.TextMatrix(i, 4)) + Val(GRDSTOCK.TextMatrix(i, 5))
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    LBLSTAOCKVALUE.Caption = N
    Screen.MousePointer = vbNormal
    Exit Sub
   
Errhand:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select DISTINCT AREA From CUSTMAST ORDER BY AREA", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!Area) Then cmbarea.AddItem (RSTCOMPANY!Area)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select ACT_CODE,ACT_NAME From CUSTMAST ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
    Do Until RSTCOMPANY.EOF
        If Not IsNull(RSTCOMPANY!ACT_NAME) Then cmbcustomer.AddItem (RSTCOMPANY!ACT_NAME)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing

    ACT_FLAG = True
    REPFLAG = True
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.value = Format(Date, "DD/MM/YYYY")
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If REPFLAG = False Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo Errhand
   
   'If Len(tXTMEDICINE.Text) < 2 Then Exit Sub
   
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        REPFLAG = False
    End If
    
    Set Me.DataList1.RowSource = RSTREP
    DataList1.ListField = "ITEM_NAME"
    DataList1.BoundColumn = "ITEM_CODE"
   
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
Errhand:
    MsgBox Err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList1.VisibleCount = 0 Then Exit Sub
            DataList1.SetFocus
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'        Case vbKeyReturn
'            If DataList1.BoundText = "" Then Exit Sub
'            LSTDISTI.SetFocus
'
'    End Select
End Sub

