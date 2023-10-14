VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRcptRegstr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT REGISTER"
   ClientHeight    =   8745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15870
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
   ScaleHeight     =   8745
   ScaleWidth      =   15870
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   8925
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   15945
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
         Left            =   10440
         TabIndex        =   5
         Top             =   825
         Width           =   1110
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
         Left            =   9135
         TabIndex        =   4
         Top             =   825
         Width           =   1110
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
         TabIndex        =   6
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
            TabIndex        =   11
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
            TabIndex        =   10
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
            TabIndex        =   9
            Top             =   255
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   7
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   8
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   7410
         Left            =   105
         TabIndex        =   3
         Top             =   1425
         Width           =   15795
         _ExtentX        =   27861
         _ExtentY        =   13070
         _Version        =   393216
         Rows            =   1
         Cols            =   22
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         FocusRect       =   2
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   8205
         TabIndex        =   12
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   30015489
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   10035
         TabIndex        =   13
         Top             =   360
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   30015489
         CurrentDate     =   40498
      End
      Begin VB.Label LBLPAIDAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   11760
         TabIndex        =   16
         Top             =   765
         Width           =   2115
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amount"
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
         Height          =   240
         Index           =   0
         Left            =   11760
         TabIndex        =   17
         Top             =   480
         Width           =   2115
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
         TabIndex        =   15
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
         TabIndex        =   14
         Top             =   405
         Width           =   285
      End
   End
End
Attribute VB_Name = "FRMRcptRegstr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CMDDISPLAY_Click()
    Call Fillgrid
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case 97, 49
                FRMESTIMATE.Show
                FRMESTIMATE.SetFocus
        End Select
    End If
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "Customer"
    GRDTranx.TextMatrix(0, 1) = "Sl"
    GRDTranx.TextMatrix(0, 2) = "Receipt Date"
    GRDTranx.TextMatrix(0, 4) = "Receipt Amt"
    GRDTranx.TextMatrix(0, 6) = "Ref."
    GRDTranx.TextMatrix(0, 7) = "Comp Ref."
    GRDTranx.TextMatrix(0, 5) = "Date"
    GRDTranx.TextMatrix(0, 20) = "Entry Date"
    
    GRDTranx.ColWidth(0) = 3000
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 1700
    GRDTranx.ColWidth(3) = 0
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 0
    GRDTranx.ColWidth(6) = 1600
'    GRDTranx.ColWidth(7) = 0
    GRDTranx.ColWidth(8) = 0
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 0
'    GRDTranx.ColWidth(15) = 0
'    GRDTranx.ColWidth(16) = 0
'    GRDTranx.ColWidth(17) = 0
'    GRDTranx.ColWidth(18) = 0
'    GRDTranx.ColWidth(19) = 0
'    GRDTranx.ColWidth(20) = 1200
'    GRDTranx.ColWidth(21) = 0
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 200
    Top = 0
    TXTDEALER.Text = " "
    TXTDEALER.Text = ""
    
    Month (Date) - 2
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.value = Format(Date, "DD/MM/YYYY")
    
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
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.Text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE (ACT_CODE <> '130000' or ACT_CODE <> '130001') And ACT_CODE Like '" & Me.TxtCode.Text & "%'ORDER BY ACT_CODE", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
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
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_CODE <> '130000' AND ACT_CODE <> '130001' And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    GRDTranx.Rows = 1
    TxtCode.Text = DataList2.BoundText
    CMDDISPLAY_Click
    DataList2.SetFocus
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "Sale Bil..."
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
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    LBLPAIDAMT.Caption = ""
    i = 1
    
    Set rstTRANX = New ADODB.Recordset
    If DataList2.BoundText = "" Then
        rstTRANX.Open "SELECT * From DBTPYMT WHERE TRX_TYPE = 'RT' AND INV_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, INV_DATE,VCH_NO", db, adOpenForwardOnly
    Else
        rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND TRX_TYPE = 'RT' AND INV_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND INV_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND (TRX_TYPE='VI')  ORDER BY TRX_TYPE, INV_DATE,VCH_NO", db, adOpenForwardOnly
    End If
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0
        
        GRDTranx.TextMatrix(i, 0) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
        GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
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
        GRDTranx.Row = i
        GRDTranx.Col = 0
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Function

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.Text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, 40
            
            If DataList2.VisibleCount = 0 Then TXTDEALER.SetFocus
            DataList2.Text = lbldealer.Caption
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
