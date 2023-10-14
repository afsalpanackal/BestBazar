VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Begin VB.Form FRMTRNXREG 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TRANSACTION REPORT"
   ClientHeight    =   10650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17475
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmTrnxReg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10650
   ScaleWidth      =   17475
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00E4CEC5&
      Height          =   10095
      Left            =   -90
      TabIndex        =   0
      Top             =   -225
      Width           =   16440
      Begin VB.ComboBox CmbPend 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "FrmTrnxReg.frx":030A
         Left            =   12255
         List            =   "FrmTrnxReg.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3180
         Visible         =   0   'False
         Width           =   3075
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         Height          =   1260
         Left            =   8445
         TabIndex        =   24
         Top             =   615
         Width           =   2715
         Begin VB.OptionButton OptPend 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Cheque Pending"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   26
            Top             =   330
            Width           =   2505
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00C0C0FF&
            Caption         =   "All"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   25
            Top             =   735
            Value           =   -1  'True
            Width           =   2505
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         Height          =   1260
         Left            =   5700
         TabIndex        =   20
         Top             =   615
         Width           =   2715
         Begin VB.OptionButton OptBoth 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Both"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   23
            Top             =   885
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.OptionButton OptRcpts 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Receipts"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   22
            Top             =   540
            Width           =   2505
         End
         Begin VB.OptionButton OptPymnt 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Payments"
            ForeColor       =   &H00000040&
            Height          =   300
            Left            =   90
            TabIndex        =   21
            Top             =   210
            Value           =   -1  'True
            Width           =   2505
         End
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
         Height          =   450
         Left            =   10095
         TabIndex        =   10
         Top             =   8535
         Visible         =   0   'False
         Width           =   1530
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
         Left            =   8730
         TabIndex        =   9
         Top             =   8535
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
         Left            =   7470
         TabIndex        =   8
         Top             =   8550
         Width           =   1200
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00E4CEC5&
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
         TabIndex        =   11
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
            BackColor       =   &H00E4CEC5&
            Caption         =   "Party Name"
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   45
            TabIndex        =   4
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00E4CEC5&
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
            Format          =   123076609
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
            Format          =   123076609
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
            TabIndex        =   12
            Top             =   300
            Width           =   285
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6615
         Left            =   150
         TabIndex        =   7
         Top             =   1875
         Width           =   16245
         _ExtentX        =   28654
         _ExtentY        =   11668
         _Version        =   393216
         Rows            =   1
         Cols            =   14
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   330
         Left            =   5730
         TabIndex        =   15
         Tag             =   "5"
         Top             =   270
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   582
         Picture         =   "FrmTrnxReg.frx":0329
         ForeColor       =   0
         BarPicture      =   "FrmTrnxReg.frx":0345
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
         Caption         =   "Total Receipts"
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
         Left            =   165
         TabIndex        =   19
         Top             =   8625
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
         Left            =   1155
         TabIndex        =   18
         Top             =   8655
         Width           =   2220
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
         Left            =   4590
         TabIndex        =   17
         Top             =   8655
         Width           =   2220
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Total Payments"
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
         TabIndex        =   16
         Top             =   8625
         Width           =   1230
      End
   End
End
Attribute VB_Name = "FRMTRNXREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmbPend_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPend.ListIndex = -1 Then Exit Sub
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM CASHATRXFILE WHERE  TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND REC_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 10) & " AND INV_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                If CmbPend.ListIndex = 0 Then
                    rststock!CHQ_STATUS = "Y"
                Else
                    rststock!CHQ_STATUS = "N"
                End If
                rststock.Update
                GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = CmbPend.text
            End If
            rststock.Close
            Set rststock = Nothing
            GRDTranx.Enabled = True
            CmbPend.Visible = False
            GRDTranx.SetFocus
        Case vbKeyEscape
            CmbPend.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub CmbPend_LostFocus()
    CmbPend.Visible = False
End Sub

Private Sub CmDDisplay_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim M As Long
    
    lblcrdt.Caption = "0.00"
    LBLTRXTOTAL.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
       rstTRANX.Open "SELECT * From CASHATRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From CASHATRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    M = 0
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        If OptRcpts.Value = True Then
            If OptPend.Value = True Then
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND INV_TYPE = 'RT' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE INV_TYPE = 'RT' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
        If OptPymnt.Value = True Then
            If OptPend.Value = True Then
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND INV_TYPE = 'PY' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE INV_TYPE = 'PY' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
        If Optboth.Value = True Then
            If OptPend.Value = True Then
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                rstTRANX.Open "SELECT * From CASHATRXFILE WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
    Else
        If OptRcpts.Value = True Then
            If OptPend.Value = True Then
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND INV_TYPE = 'RT' AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE INV_TYPE = 'RT' AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
        If OptPymnt.Value = True Then
            If OptPend.Value = True Then
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND INV_TYPE = 'PY' AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE INV_TYPE = 'PY' AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
        If Optboth.Value = True Then
            If OptPend.Value = True Then
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE CASH_MODE ='B' AND CHQ_STATUS = 'N' AND ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            Else
                If OptRcpts.Value = True Then rstTRANX.Open "SELECT * From CASHATRXFILE WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY VCH_DATE,REC_NO", db, adOpenStatic, adLockReadOnly
            End If
        End If
    End If
    
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        Select Case rstTRANX!INV_TYPE
            Case "RT"
                GRDTranx.TextMatrix(M, 2) = Format(Round(rstTRANX!AMOUNT, 2), "0.00")
            Case "PY"
                GRDTranx.TextMatrix(M, 3) = Format(Round(rstTRANX!AMOUNT, 2), "0.00")
        End Select

        Select Case rstTRANX!CASH_MODE
            Case "C"
                GRDTranx.TextMatrix(M, 4) = "By Cash"
                GRDTranx.TextMatrix(M, 5) = ""
                GRDTranx.TextMatrix(M, 6) = ""
                GRDTranx.TextMatrix(M, 7) = ""
                GRDTranx.TextMatrix(M, 8) = ""
            Case "B"
                GRDTranx.TextMatrix(M, 4) = "To Bank"
                GRDTranx.TextMatrix(M, 5) = IIf(IsNull(rstTRANX!CHQ_NO), "", rstTRANX!CHQ_NO)
                GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!CHQ_DATE), "", rstTRANX!CHQ_DATE)
                GRDTranx.TextMatrix(M, 7) = IIf(IsNull(rstTRANX!BANK), "", rstTRANX!BANK)
                GRDTranx.TextMatrix(M, 8) = IIf((rstTRANX!CHQ_STATUS = "Y"), "Passed", "Pending")
        End Select
        GRDTranx.TextMatrix(M, 9) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(M, 10) = rstTRANX!REC_NO
        GRDTranx.TextMatrix(M, 11) = rstTRANX!INV_TYPE
        GRDTranx.TextMatrix(M, 12) = rstTRANX!INV_NO
        GRDTranx.TextMatrix(M, 13) = IIf(IsNull(rstTRANX!BILL_TRX_TYPE), "", rstTRANX!BILL_TRX_TYPE)
              
        lblcrdt.Caption = Format(Val(lblcrdt.Caption) + Val(GRDTranx.TextMatrix(M, 2)), "0.00")
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + Val(GRDTranx.TextMatrix(M, 3)), "0.00")
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
        
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    GRDTranx.SetFocus
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
''    Screen.MousePointer = vbHourglass
''    ReportNameVar = Rptpath & "RPTPURCHREG"
''    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
''    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
''    Set CRXFormulaFields = Report.FormulaFields
''    For i = 1 To Report.Database.Tables.Count
''        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
''    Next i
''    For Each CRXFormulaField In CRXFormulaFields
''        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
''    Next
''    frmreport.Caption = "PURCHASE REGISTER"
''    Call GENERATEREPORT
''    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "PARTY NAME"
    GRDTranx.TextMatrix(0, 2) = "RECEIPT"
    GRDTranx.TextMatrix(0, 3) = "PAYMENT"
    GRDTranx.TextMatrix(0, 4) = "Cash Mode"
    GRDTranx.TextMatrix(0, 5) = "Cheque No"
    GRDTranx.TextMatrix(0, 6) = "Chq Date"
    GRDTranx.TextMatrix(0, 7) = "Bank"
    GRDTranx.TextMatrix(0, 8) = "Status"
    GRDTranx.TextMatrix(0, 9) = "TRX_TYPE"
    GRDTranx.TextMatrix(0, 10) = "REC_NO"
    GRDTranx.TextMatrix(0, 11) = "INV_TYPE"
    GRDTranx.TextMatrix(0, 12) = "INV_NO"
    GRDTranx.TextMatrix(0, 13) = "BILL_TRX_TYPE"
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 3000
    GRDTranx.ColWidth(2) = 1400
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1800
    GRDTranx.ColWidth(6) = 1400
    GRDTranx.ColWidth(7) = 2500
    GRDTranx.ColWidth(8) = 1500
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 7
    GRDTranx.ColAlignment(3) = 7
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 1
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(7) = 1
    GRDTranx.ColAlignment(8) = 4
        
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    Me.Width = 16425
    Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
    ACT_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.TextMatrix(GRDTranx.Row, 4) <> "To Bank" Then Exit Sub
            Select Case GRDTranx.Col
                 Case 8
                    CmbPend.Visible = True
                    CmbPend.Top = GRDTranx.CellTop + 1900
                    CmbPend.Left = GRDTranx.CellLeft + 150
                    CmbPend.Width = GRDTranx.CellWidth
                    'CmbPend.Height = GRDTranx.CellHeight
                    If Trim(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = "Passed" Then
                        CmbPend.ListIndex = 0
                    Else
                        CmbPend.ListIndex = 1
                   End If
                CmbPend.SetFocus
            End Select
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

Private Sub TXTDEALER_GotFocus()
    OPTCUSTOMER.Value = True
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
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
        If OptPymnt.Value = True Then
            If ACT_FLAG = True Then
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311' And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
                ACT_FLAG = False
            Else
                ACT_REC.Close
                ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311' And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
                ACT_FLAG = False
            End If
        Else
            If ACT_FLAG = True Then
                ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
                ACT_FLAG = False
            Else
                ACT_REC.Close
                ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST  WHERE ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
                ACT_FLAG = False
            End If
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

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    LBLTRXTOTAL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "EzBiz"
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
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    
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
        Print #1, Space(52) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(52) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(52) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "Purchase Register for the Period"
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & "From " & DTFROM.Value & " TO " & DTTo.Value
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
    RSTTRXFILE.Open "SELECT * From SALESREG ORDER BY VCH_DATE", db, adOpenStatic, adLockReadOnly
    Do Until RSTTRXFILE.EOF
        SN = SN + 1
        Print #1, Chr(27) & Chr(71) & Space(9) & AlignRight(str(SN), 3) & "." & Space(1) & _
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
    'MsgBox "Report file generated at " & Rptpath & "Report.PRN" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub CMDDISPLAY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            DTTo.SetFocus
    End Select
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
