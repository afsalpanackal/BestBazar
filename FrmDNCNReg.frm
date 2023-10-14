VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDNCNREG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT REGISTER"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18645
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDNCNReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8745
   ScaleWidth      =   18645
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8925
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   18735
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   8295
         TabIndex        =   23
         Top             =   825
         Width           =   1200
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
         Height          =   555
         Left            =   9525
         TabIndex        =   4
         Top             =   825
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
         Height          =   555
         Left            =   6945
         TabIndex        =   3
         Top             =   825
         Width           =   1320
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
         Height          =   1305
         Left            =   120
         TabIndex        =   5
         Top             =   105
         Width           =   6840
         Begin VB.TextBox TxtCode 
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   1095
            TabIndex        =   10
            Top             =   210
            Width           =   1875
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
            Height          =   360
            Left            =   3030
            TabIndex        =   1
            Top             =   210
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   3030
            TabIndex        =   2
            Top             =   585
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
         Begin VB.PictureBox rptPRINT 
            Height          =   480
            Left            =   9990
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   9
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "CUSTOMER"
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
            Index           =   5
            Left            =   45
            TabIndex        =   8
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   6
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   7
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   8205
         TabIndex        =   11
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
         Left            =   10035
         TabIndex        =   12
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   123076609
         CurrentDate     =   40498
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F4F0DB&
         Height          =   570
         Left            =   13980
         TabIndex        =   19
         Top             =   795
         Width           =   4680
         Begin VB.OptionButton Optboth 
            BackColor       =   &H00F4F0DB&
            Caption         =   "Both"
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
            Height          =   315
            Left            =   3045
            TabIndex        =   22
            Top             =   180
            Width           =   1560
         End
         Begin VB.OptionButton OptcrNote 
            BackColor       =   &H00F4F0DB&
            Caption         =   "Credit Note"
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
            Height          =   315
            Left            =   90
            TabIndex        =   21
            Top             =   180
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.OptionButton OptDnNote 
            BackColor       =   &H00F4F0DB&
            Caption         =   "Debit Note"
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
            Height          =   315
            Left            =   1530
            TabIndex        =   20
            Top             =   180
            Width           =   1560
         End
      End
      Begin VB.Frame Frame1 
         Height          =   6705
         Left            =   105
         TabIndex        =   16
         Top             =   1305
         Width           =   18600
         Begin MSFlexGridLib.MSFlexGrid GRDTranx 
            Height          =   6570
            Left            =   15
            TabIndex        =   17
            Top             =   120
            Width           =   18585
            _ExtentX        =   32782
            _ExtentY        =   11589
            _Version        =   393216
            Rows            =   1
            Cols            =   24
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            BackColorBkg    =   12632256
            FocusRect       =   2
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
      End
      Begin VB.Label lbladdress 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   13965
         TabIndex        =   18
         Top             =   210
         Width           =   4680
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblagent 
         Caption         =   " "
         Height          =   270
         Left            =   11595
         TabIndex        =   15
         Top             =   360
         Visible         =   0   'False
         Width           =   2310
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
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
         Index           =   10
         Left            =   7035
         TabIndex        =   14
         Top             =   405
         Width           =   1140
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
         Index           =   9
         Left            =   9750
         TabIndex        =   13
         Top             =   405
         Width           =   285
      End
   End
End
Attribute VB_Name = "FRMDNCNREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub



Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "DATE"
    GRDTranx.TextMatrix(0, 3) = "NO."
    GRDTranx.TextMatrix(0, 4) = "Debit"
    GRDTranx.TextMatrix(0, 5) = "Credit"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "" '"CR NO"
    GRDTranx.TextMatrix(0, 8) = "" '"TYPE"
    GRDTranx.TextMatrix(0, 20) = "Entry Date"
    GRDTranx.TextMatrix(0, 21) = "Bank Name"
    GRDTranx.TextMatrix(0, 22) = "Customer"
    GRDTranx.TextMatrix(0, 23) = "GSTIN No."
    
    GRDTranx.ColWidth(0) = 900
    GRDTranx.ColWidth(1) = 700
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 1200
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1200
    GRDTranx.ColWidth(6) = 1400
    GRDTranx.ColWidth(7) = 0
    GRDTranx.ColWidth(8) = 0
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
    GRDTranx.ColWidth(15) = 0
    GRDTranx.ColWidth(16) = 0
    GRDTranx.ColWidth(17) = 0
    GRDTranx.ColWidth(18) = 0
    GRDTranx.ColWidth(19) = 0
    GRDTranx.ColWidth(20) = 1100
    GRDTranx.ColWidth(21) = 2000
    GRDTranx.ColWidth(22) = 3000
    GRDTranx.ColWidth(23) = 2000
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    'GRDTranx.ColAlignment(4) = 4
    'GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(21) = 1
    GRDTranx.ColAlignment(22) = 1
    GRDTranx.ColAlignment(23) = 1
    'GRDTranx.ColAlignment(28) = 1
    
    
    'GRDBILL.ColAlignment(6) = 4
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 200
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.MNUPYMNT.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub



Private Sub TXTCODE_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
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
    TxtCode.text = DataList2.BoundText
    CmDDisplay_Click
    DataList2.SetFocus
    flagchange.Caption = ""
    lbldealer.Caption = DataList2.text
    
    On Error GoTo ERRHAND
    Dim rstCustomer As ADODB.Recordset
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "select * from CUSTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstCustomer.EOF And rstCustomer.BOF) Then
        lbladdress.Caption = IIf(IsNull(rstCustomer!Address), "", Trim(rstCustomer!Address))
    Else
        lbladdress.Caption = ""
    End If
        
    'TXTDEALER.Text = lbldealer.Caption
    'LBL.Caption = ""
    Exit Sub
ERRHAND:
    
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
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
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim RSTACTMAST As ADODB.Recordset
    Dim RSTBANK As ADODB.Recordset
    Dim i As Long
        
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    GRDTranx.rows = 1
    
    i = 1
    Set rstTRANX = New ADODB.Recordset
    If DataList2.BoundText = "" Then
        If OptcrNote.Value = True Then
            rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'CB') AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
            'rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'CB') ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        ElseIf OptDnNote.Value = True Then
            rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'DB' ) AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        Else
            rstTRANX.Open "SELECT * From DBTPYMT WHERE (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' ) AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        End If
    Else
        If OptcrNote.Value = True Then
            rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'CB') AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        ElseIf OptDnNote.Value = True Then
            rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'DB' ) AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        Else
            rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' ) AND INV_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
        End If
    End If
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!REC_NO), "", rstTRANX!REC_NO)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
        GRDTranx.TextMatrix(i, 22) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        Set RSTACTMAST = New ADODB.Recordset
        RSTACTMAST.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTACTMAST.EOF And RSTACTMAST.BOF) Then
            GRDTranx.TextMatrix(i, 23) = IIf(IsNull(RSTACTMAST!KGST), "", RSTACTMAST!KGST)
        End If
        RSTACTMAST.Close
        Set RSTACTMAST = Nothing
        
        Select Case rstTRANX!check_flag
            Case "Y"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!INV_AMT, "0.00")
            Case "N"
                GRDTranx.TextMatrix(i, 5) = "" '0 '""Format(rstTRANX!INV_AMT, "0.00")
        End Select
        Select Case rstTRANX!TRX_TYPE
           
            Case "DB"
                GRDTranx.TextMatrix(i, 0) = "Debit Note"
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbRed
            Case "CB"
                GRDTranx.TextMatrix(i, 0) = "Credit Note"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                
                GRDTranx.CellForeColor = vbBlue
            Case "SR"
                GRDTranx.TextMatrix(i, 0) = "SALES RETURN"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "RW"
                GRDTranx.TextMatrix(i, 0) = "SALES RETURN(W)"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
        End Select
               
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!INV_TRX_TYPE), "", rstTRANX!INV_TRX_TYPE)
        Select Case rstTRANX!BANK_FLAG
            Case "Y"
                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_TYPE), "", rstTRANX!B_TRX_TYPE)
                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
                GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!B_BILL_TRX_TYPE), "", rstTRANX!B_BILL_TRX_TYPE)
                GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!B_TRX_YEAR), "", rstTRANX!B_TRX_YEAR)
                GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
                
                Set RSTBANK = New ADODB.Recordset
                RSTBANK.Open "select * from BANKCODE  WHERE BANK_CODE = '" & GRDTranx.TextMatrix(i, 13) & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTBANK.EOF And RSTBANK.BOF) Then
                    GRDTranx.TextMatrix(i, 21) = IIf(IsNull(RSTBANK!BANK_NAME), "", RSTBANK!BANK_NAME)
                End If
                RSTBANK.Close
                Set RSTBANK = Nothing
                
                GRDTranx.TextMatrix(i, 15) = ""
                GRDTranx.TextMatrix(i, 16) = ""
                GRDTranx.TextMatrix(i, 17) = ""
                GRDTranx.TextMatrix(i, 18) = ""
                GRDTranx.TextMatrix(i, 19) = ""
            Case Else
                GRDTranx.TextMatrix(i, 9) = ""
                GRDTranx.TextMatrix(i, 10) = ""
                GRDTranx.TextMatrix(i, 11) = ""
                GRDTranx.TextMatrix(i, 12) = ""
                GRDTranx.TextMatrix(i, 13) = ""
                GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstTRANX!C_TRX_TYPE), "", rstTRANX!C_TRX_TYPE)
                GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rstTRANX!C_REC_NO), "", rstTRANX!C_REC_NO)
                GRDTranx.TextMatrix(i, 17) = IIf(IsNull(rstTRANX!C_INV_TRX_TYPE), "", rstTRANX!C_INV_TRX_TYPE)
                GRDTranx.TextMatrix(i, 18) = IIf(IsNull(rstTRANX!C_INV_TYPE), "", rstTRANX!C_INV_TYPE)
                GRDTranx.TextMatrix(i, 19) = IIf(IsNull(rstTRANX!C_INV_NO), "", rstTRANX!C_INV_NO)
        End Select
        GRDTranx.TextMatrix(i, 20) = IIf(IsNull(rstTRANX!ENTRY_DATE), "", rstTRANX!ENTRY_DATE)
        GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    If GRDTranx.rows > 16 Then GRDTranx.TopRow = GRDTranx.rows - 1
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            DataList2.text = lbldealer.Caption
            Call DataList2_Click
            'lbladdress.Caption = ""
            DataList2.SetFocus
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Sale = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset

    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    ReportNameVar = Rptpath & "RptDrCr"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    If DataList2.BoundText = "" Then
        If OptcrNote.Value = True Then
            Report.RecordSelectionFormula = "({DBTPYMT.TRX_TYPE} ='CB' AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        ElseIf OptDnNote.Value = True Then
            Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='DB' ) AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        Else
            Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' ) AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        End If
    Else
        If OptcrNote.Value = True Then
            Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "') and {DBTPYMT.TRX_TYPE} ='CB' ) AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        ElseIf OptDnNote.Value = True Then
            Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "') and {DBTPYMT.TRX_TYPE} ='DB' ) AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        Else
            Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "') and {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' ) AND ({DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
        End If
    End If
    
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
        If OptcrNote.Value = True Then
            If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Credit Note Statement for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        ElseIf OptDnNote.Value = True Then
            If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Debit Note Statement for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        Else
            If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Credit And Debit Note Statement for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        End If
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub
