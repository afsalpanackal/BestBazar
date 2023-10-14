VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMPaymntreg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT REGISTER"
   ClientHeight    =   9705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
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
   Icon            =   "FrmpPaymentreg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9705
   ScaleWidth      =   9495
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      Caption         =   "PRESS ESC TO CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4830
      Left            =   1140
      TabIndex        =   6
      Top             =   1755
      Visible         =   0   'False
      Width           =   7005
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   3360
         Left            =   105
         TabIndex        =   7
         Top             =   1335
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   5927
         _Version        =   393216
         Rows            =   1
         Cols            =   6
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
      Begin VB.Label LBLPAID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3615
         TabIndex        =   30
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "PAID AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   29
         Top             =   735
         Width           =   930
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "BAL AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   8
         Left            =   4935
         TabIndex        =   28
         Top             =   735
         Width           =   870
      End
      Begin VB.Label LBLBAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4815
         TabIndex        =   27
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLINVDATE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   45
         TabIndex        =   18
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV DATE"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   315
         TabIndex        =   17
         Top             =   735
         Width           =   885
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1080
         TabIndex        =   16
         Top             =   315
         Width           =   4125
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   900
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2445
         TabIndex        =   11
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   2595
         TabIndex        =   10
         Top             =   735
         Width           =   810
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV NO"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   9
         Top             =   735
         Width           =   675
      End
      Begin VB.Label LBLBILLNO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1425
         TabIndex        =   8
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0C0FF&
      Height          =   9870
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   9570
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTAL"
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   795
         Left            =   150
         TabIndex        =   20
         Top             =   8925
         Width           =   9375
         Begin VB.Label LBLPAIDAMT 
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
            ForeColor       =   &H00000080&
            Height          =   420
            Left            =   4440
            TabIndex        =   26
            Top             =   225
            Width           =   1845
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "PAID AMT"
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
            Index           =   3
            Left            =   3435
            TabIndex        =   25
            Top             =   300
            Width           =   1080
         End
         Begin VB.Label LBLINVAMT 
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
            ForeColor       =   &H00000080&
            Height          =   420
            Left            =   1455
            TabIndex        =   24
            Top             =   225
            Width           =   1890
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "INVOICE AMT"
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
            Left            =   195
            TabIndex        =   23
            Top             =   300
            Width           =   1395
         End
         Begin VB.Label LBLBALAMT 
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
            ForeColor       =   &H00000080&
            Height          =   420
            Left            =   7230
            TabIndex        =   22
            Top             =   225
            Width           =   2040
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "BAL AMT"
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
            Left            =   6345
            TabIndex        =   21
            Top             =   300
            Width           =   930
         End
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
         Height          =   480
         Left            =   6795
         TabIndex        =   5
         Top             =   900
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
         Left            =   6795
         TabIndex        =   4
         Top             =   345
         Width           =   1440
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
         Height          =   1320
         Left            =   135
         TabIndex        =   12
         Top             =   105
         Width           =   5670
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
            Left            =   1485
            TabIndex        =   1
            Top             =   225
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1485
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
            Height          =   315
            Index           =   5
            Left            =   300
            TabIndex        =   19
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   13
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   14
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   7380
         Left            =   105
         TabIndex        =   3
         Top             =   1470
         Width           =   9420
         _ExtentX        =   16616
         _ExtentY        =   13018
         _Version        =   393216
         Rows            =   1
         Cols            =   9
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "FRMPaymntreg"
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
            CMDDISPLAY.SetFocus
    End Select
End Sub

Private Sub CMDDISPLAY_Click()
    Call FILLGRID
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    Call FILLGRID
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "STATUS"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "INV DATE"
    GRDTranx.TextMatrix(0, 3) = "INV NO"
    GRDTranx.TextMatrix(0, 4) = "INV AMT"
    GRDTranx.TextMatrix(0, 5) = "PAID AMT"
    GRDTranx.TextMatrix(0, 6) = "BAL AMT"
    GRDTranx.TextMatrix(0, 7) = "TYPE"
    GRDTranx.TextMatrix(0, 8) = "DAYS"
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 600
    GRDTranx.ColWidth(2) = 1300
    GRDTranx.ColWidth(3) = 1100
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 1100
    GRDTranx.ColWidth(6) = 1100
    GRDTranx.ColWidth(7) = 1200
    GRDTranx.ColWidth(8) = 700
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 3
    GRDTranx.ColAlignment(6) = 3
    GRDTranx.ColAlignment(7) = 3
    GRDTranx.ColAlignment(8) = 3
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Pymnt Date"
    GRDBILL.TextMatrix(0, 2) = "Paid Amt"
    GRDBILL.TextMatrix(0, 3) = "Pymnt No"
    GRDBILL.TextMatrix(0, 4) = "Entry Date"
    GRDBILL.TextMatrix(0, 5) = "Ref No"
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 1200
    GRDBILL.ColWidth(2) = 1200
    GRDBILL.ColWidth(3) = 1200
    GRDBILL.ColWidth(4) = 1200
    GRDBILL.ColWidth(5) = 1200

    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(1) = 3
    GRDBILL.ColAlignment(2) = 3
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    
    CLOSEALL = 1
    ACT_FLAG = True
    Me.Width = 9585
    Me.Height = 10185
    Me.Left = 1500
    Me.Top = 0
    
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
            LBLSUPPLIER.Caption = " " & DataList2.Text
            LBLINVDATE.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 2)
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 4)
            LBLPAID.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 5)
            LBLBAL.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)
            
            GRDBILL.Rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND INV_NO =  " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'PY'", db2, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!RCPT_NO
                GRDBILL.TextMatrix(i, 1) = Format(RSTTRXFILE!RCPT_DATE, "DD/MM/YYYY")
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!RCPT_AMOUNT, "0.00")
                GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!RCPT_ENTRY_DATE
                GRDBILL.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
        Case vbKeyF6
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "PEND" Then Exit Sub
            FRMPaymntreg.Enabled = False
            FRMPAYMENTSHORT.LBLSUPPLIER.Caption = DataList2.Text
            FRMPAYMENTSHORT.lblactcode.Caption = DataList2.BoundText
            FRMPAYMENTSHORT.LBLINVDATE.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            FRMPAYMENTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            FRMPAYMENTSHORT.LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            FRMPAYMENTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            FRMPAYMENTSHORT.LBLBALAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            'FRMPAYMENTSHORT.LBLTYPE.Caption = Trim(GRDTranx.TextMatrix(GRDTranx.Row, 7))
            FRMPAYMENTSHORT.Show
    End Select
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
    'LBL.Caption = ""
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

Private Function FILLGRID()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Integer
    
    
    If DataList2.BoundText = "" Then Exit Function
   ' On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    i = 1
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From CRDTPYMT WHERE [ACT_CODE] = '" & DataList2.BoundText & "' ORDER BY INV_NO DESC", db2, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = rstTRANX!INV_NO
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
        GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
        GRDTranx.TextMatrix(i, 6) = Format(rstTRANX!INV_AMT - rstTRANX!RCPT_AMT, "0.00")
        GRDTranx.TextMatrix(i, 7) = rstTRANX!TRX_TYPE
        If rstTRANX!CHECK_FLAG = "Y" Then
            GRDTranx.TextMatrix(i, 0) = "PAID"
        Else
             GRDTranx.TextMatrix(i, 0) = "PEND"
             GRDTranx.TextMatrix(i, 8) = DateDiff("d", GRDTranx.TextMatrix(i, 2), Date)
        End If
        GRDTranx.Row = i
        GRDTranx.Col = 0
        If rstTRANX!CHECK_FLAG = "N" Then
            LBLBALAMT.Caption = Format(Val(LBLBALAMT.Caption) + Val(GRDTranx.TextMatrix(i, 6)), "0.00")
            GRDTranx.CellForeColor = vbRed
        Else
            GRDTranx.CellForeColor = vbBlue
        End If
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + rstTRANX!INV_AMT, "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + rstTRANX!RCPT_AMT, "0.00")
        LBLBALAMT.Caption = Format(Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption), "0.00")
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
