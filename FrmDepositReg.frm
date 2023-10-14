VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDepositReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Savings / Deposit Register"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9900
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
   Moveable        =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   9900
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00F4F0DB&
      Height          =   8595
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   9990
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00F4F0DB&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   120
         TabIndex        =   10
         Top             =   7635
         Width           =   9810
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Op. Balance"
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
            Height          =   300
            Index           =   8
            Left            =   1035
            TabIndex        =   19
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label lblOPBal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   990
            TabIndex        =   18
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLPAIDAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5325
            TabIndex        =   16
            Top             =   435
            Width           =   1875
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paid Amt"
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
            Index           =   3
            Left            =   5340
            TabIndex        =   15
            Top             =   150
            Width           =   1875
         End
         Begin VB.Label LBLINVAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   3135
            TabIndex        =   14
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd Amt"
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
            Height          =   300
            Index           =   6
            Left            =   3180
            TabIndex        =   13
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label LBLBALAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   7395
            TabIndex        =   12
            Top             =   435
            Visible         =   0   'False
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Amt"
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
            Height          =   315
            Index           =   7
            Left            =   7395
            TabIndex        =   11
            Top             =   150
            Visible         =   0   'False
            Width           =   1815
         End
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
         Height          =   480
         Left            =   5505
         TabIndex        =   5
         Top             =   825
         Width           =   1440
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
         Height          =   480
         Left            =   5505
         TabIndex        =   4
         Top             =   255
         Width           =   1440
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00F4F0DB&
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
         Height          =   1320
         Left            =   120
         TabIndex        =   6
         Top             =   45
         Width           =   5325
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
            Left            =   60
            TabIndex        =   22
            Top             =   1170
            Visible         =   0   'False
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
            Left            =   1545
            TabIndex        =   1
            Top             =   210
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1545
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
            TabIndex        =   17
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Account Name"
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
            Left            =   75
            TabIndex        =   9
            Top             =   255
            Width           =   1410
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
         Height          =   6240
         Left            =   120
         TabIndex        =   3
         Top             =   1365
         Width           =   9840
         _ExtentX        =   17357
         _ExtentY        =   11007
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         FocusRect       =   2
         SelectionMode   =   1
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F7 to make Payments"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Index           =   0
         Left            =   7140
         TabIndex        =   21
         Top             =   1020
         Width           =   3120
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 to make Receipts"
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
         Height          =   300
         Index           =   8
         Left            =   7140
         TabIndex        =   20
         Top             =   720
         Width           =   3120
      End
   End
End
Attribute VB_Name = "FRMDepositReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub CMDDISPLAY_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "INV / PAID DATE"
    GRDTranx.TextMatrix(0, 3) = "INV NO"
    GRDTranx.TextMatrix(0, 4) = "INV AMT"
    GRDTranx.TextMatrix(0, 5) = "PAID AMT"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "CR NO"
    
    GRDTranx.ColWidth(0) = 1200
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 1700
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 1600
    GRDTranx.ColWidth(7) = 0
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    
    
    CLOSEALL = 1
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 1500
    Top = 0
    TXTDEALER.Text = " "
    TXTDEALER.Text = ""
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.MNUPYMNT.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub txtPassword_GotFocus()
    txtpassword.Text = ""
    txtpassword.PasswordChar = " "
    txtpassword.SelStart = 0
    txtpassword.SelLength = Len(txtpassword.Text)
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            CMBMONTH.SetFocus
    End Select
End Sub

Private Sub TXTPASSWORD_LostFocus()
    If UCase(txtpassword.Text) = "SARAKALAM" Then
        txtpassword = "YEAR " & Year(Date)
        txtpassword.PasswordChar = ""
        CMBMONTH.SetFocus
        Height = 6000
    Else
        txtpassword = "YEAR " & Year(Date)
        txtpassword.PasswordChar = ""
        CMBMONTH.SetFocus
        Height = 3700
    End If
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            frmemain.Enabled = True
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

    Dim i As Long
    Dim E_TABLE As String
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyF6
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMDPRcpt.LBLSUPPLIER.Caption = DataList2.Text
            FRMDPRcpt.lblactcode.Caption = DataList2.BoundText
            'FRMRECEIPTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMRECEIPTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMRECEIPTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMRECEIPTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMRECEIPTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDPRcpt.Show
        Case vbKeyF7
            Enabled = False
            FRMDPPYMNT.LBLSUPPLIER.Caption = DataList2.Text
            FRMDPPYMNT.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDPPYMNT.Show
    End Select
End Sub

Private Sub GRDTranx_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHand
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
            If GRDTranx.Rows = 1 Then Exit Sub
            If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                    Select Case GRDTranx.TextMatrix(GRDTranx.Row, 0)
                        Case "Payment"
                            db.BeginTrans
                            db.Execute "delete From DBTPYMT WHERE TRX_TYPE='PL' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & ""
                            db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'PL'"
                            db.CommitTrans
                        Case "Receipt"
                            db.BeginTrans
                            db.Execute "delete From DBTPYMT WHERE TRX_TYPE='SV' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & ""
                            db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'SV'"
                            db.CommitTrans
                    End Select
                    Call Fillgrid
                Else
                    GRDTranx.SetFocus
                End If
            End If
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = -2147168237 Then
        On Error Resume Next
        db.RollbackTrans
    Else
        MsgBox Err.Description
    End If
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
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='711')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='711')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!act_name
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    GRDTranx.Rows = 1
    TxtCode.Text = DataList2.BoundText
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
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
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Long
    
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'SV' OR TRX_TYPE = 'PL') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!RCPT_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        Select Case rstTRANX!TRX_TYPE
            Case "SV"
                GRDTranx.TextMatrix(i, 0) = "Receipt"
                GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbRed
            Case "PL"
                GRDTranx.TextMatrix(i, 0) = "Payment"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.CellForeColor = vbBlue
        End Select
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        GRDTranx.Row = i
        GRDTranx.Col = 0
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLBALAMT.Caption = Format(Val(lblOPBal.Caption) + (Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption)), "0.00")
    
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ErrHand:
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
            'lbladdress.Caption = ""
            DataList2.SetFocus
        Case vbKeyEscape
            txtBillNo.Visible = True
            txtBillNo.SetFocus
    End Select
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub
