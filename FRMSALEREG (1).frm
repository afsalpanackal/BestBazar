VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSALEREG 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES RETURN REGISTER"
   ClientHeight    =   8130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FRMSALEREG.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   11550
   Begin VB.Frame FRMEBILL 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "PRESS ESC TO CANCEL"
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   1815
      TabIndex        =   22
      Top             =   2130
      Visible         =   0   'False
      Width           =   8160
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   3765
         Left            =   45
         TabIndex        =   23
         Top             =   645
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   6641
         _Version        =   393216
         Rows            =   1
         Cols            =   7
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
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
      Begin VB.Label LBLBILLNO 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   4950
         TabIndex        =   29
         Top             =   270
         Width           =   990
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
         TabIndex        =   28
         Top             =   270
         Width           =   825
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
         TabIndex        =   27
         Top             =   300
         Width           =   885
      End
      Begin VB.Label LBLBILLAMT 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   6915
         TabIndex        =   26
         Top             =   270
         Width           =   1155
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
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   25
         Top             =   285
         Width           =   990
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   1080
         TabIndex        =   24
         Top             =   270
         Width           =   3030
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFFF&
      Height          =   8295
      Left            =   -135
      TabIndex        =   0
      Top             =   -180
      Width           =   11670
      Begin VB.Frame Frame2 
         BackColor       =   &H0093DDF9&
         Height          =   1770
         Left            =   5700
         TabIndex        =   30
         Top             =   105
         Width           =   3150
         Begin VB.OptionButton OptPurchase 
            BackColor       =   &H0093DDF9&
            Caption         =   "Sales Return"
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
            Height          =   330
            Left            =   180
            TabIndex        =   33
            Top             =   840
            Value           =   -1  'True
            Width           =   2325
         End
         Begin VB.OptionButton OptPetty 
            BackColor       =   &H0093DDF9&
            Caption         =   "Petty Sales Return"
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
            Height          =   330
            Left            =   180
            TabIndex        =   32
            Top             =   420
            Width           =   2850
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H0093DDF9&
            Caption         =   "All"
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
            Height          =   330
            Left            =   180
            TabIndex        =   31
            Top             =   1230
            Width           =   1635
         End
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFFF&
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
         TabIndex        =   5
         Top             =   105
         Width           =   5550
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00C0FFFF&
            Caption         =   "PERIOD FROM"
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   45
            TabIndex        =   8
            Top             =   315
            Value           =   -1  'True
            Width           =   1620
         End
         Begin VB.OptionButton OPTCUSTOMER 
            BackColor       =   &H00C0FFFF&
            Caption         =   "CUSTOMER"
            ForeColor       =   &H00000040&
            Height          =   210
            Left            =   45
            TabIndex        =   7
            Top             =   720
            Width           =   1320
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
            Left            =   1680
            TabIndex        =   6
            Top             =   720
            Width           =   3735
         End
         Begin MSComCtl2.DTPicker DTFROM 
            Height          =   390
            Left            =   1680
            TabIndex        =   9
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
            TabIndex        =   10
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
            TabIndex        =   11
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
            TabIndex        =   14
            Top             =   300
            Width           =   285
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   13
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   12
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
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
         TabIndex        =   4
         Top             =   7755
         Width           =   1200
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
         TabIndex        =   3
         Top             =   7755
         Width           =   1290
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
         TabIndex        =   2
         Top             =   7755
         Visible         =   0   'False
         Width           =   1320
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
         Left            =   8730
         TabIndex        =   1
         Top             =   7755
         Width           =   1530
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   5295
         Left            =   150
         TabIndex        =   15
         Top             =   1890
         Width           =   11430
         _ExtentX        =   20161
         _ExtentY        =   9340
         _Version        =   393216
         Rows            =   1
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
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
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Amount"
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
         TabIndex        =   21
         Top             =   7260
         Width           =   1230
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
         TabIndex        =   20
         Top             =   7305
         Width           =   2220
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Amount"
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
         TabIndex        =   19
         Top             =   7215
         Width           =   1050
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
         TabIndex        =   18
         Top             =   7245
         Width           =   2220
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
         TabIndex        =   17
         Top             =   7290
         Width           =   2220
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
         ForeColor       =   &H000000FF&
         Height          =   480
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   7260
         Width           =   1020
      End
   End
End
Attribute VB_Name = "FRMSALEREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CmDDisplay_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim n, M As Long
    Dim TaxAmt, EXSALEAMT, TAXSALEAMT, MRPVALUE As Double
    Dim TAXRATE As Single
    
    db.Execute "delete From SALESREG"
    
    lblcrdt.Caption = "0.00"
    lblcash.Caption = "0.00"
    LBLTRXTOTAL.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.rows = 1
    'vbalProgressBar1.value = 0
    'vbalProgressBar1.ShowText = True
    
    n = 1
    M = 0
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
'    Set rstTRANX = New ADODB.Recordset
'    If OPTPERIOD.value = True Then
'        rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'    Else
'        rstTRANX.Open "SELECT * From RETURNMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'    End If
'
'    If rstTRANX.RecordCount > 0 Then
'        vbalProgressBar1.Max = rstTRANX.RecordCount
'    Else
'        vbalProgressBar1.Max = 100
'    End If
'    rstTRANX.Close
'    Set rstTRANX = Nothing
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.Value = True Then
        If OptPurchase.Value = True Then
            rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptPetty.Value = True Then
            rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RW' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR' OR TRX_TYPE='RW') ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    Else
        If OptPurchase.Value = True Then
            rstTRANX.Open "SELECT * From RETURNMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        ElseIf OptPetty.Value = True Then
            rstTRANX.Open "SELECT * From RETURNMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND TRX_TYPE='RW' ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        Else
            rstTRANX.Open "SELECT * From RETURNMAST WHERE ACT_CODE = '" & DataList2.BoundText & "' AND VCH_DATE <= '" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' AND (TRX_TYPE='SR' OR TRX_TYPE='RW') ORDER BY VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
        End If
    End If
    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 7) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(M, 2) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 3) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
        GRDTranx.TextMatrix(M, 4) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 5) = Format(Round(Val(GRDTranx.TextMatrix(M, 3)) - Val(GRDTranx.TextMatrix(M, 4)), 2), "0.00")
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        
        CMDDISPLAY.Tag = ""
        FRMEMAIN.Tag = ""
        FRMEBILL.Tag = ""
        
        'If rstTRANX!TRX_TYPE <> "SR" Then GoTo SKIP
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select DISTINCT SALES_TAX From RTRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            EXSALEAMT = 0
            TAXSALEAMT = 0
            TaxAmt = 0
            MRPVALUE = 0
            RSTSALEREG.AddNew
            RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
            RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
            RSTSALEREG!DISCOUNT = Val(GRDTranx.TextMatrix(M, 4))
            RSTSALEREG!VCH_AMOUNT = Val(GRDTranx.TextMatrix(M, 3))
            RSTSALEREG!PAYAMOUNT = 0
            RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
            RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
            RSTSALEREG.Update
            
            RSTTRXFILE.MoveNext
        Loop
    
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        If (rstTRANX!POST_FLAG = "Y") Then
            lblcash.Caption = Format(Val(lblcash.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        Else
            lblcrdt.Caption = Format(Val(lblcrdt.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        End If
        'vbalProgressBar1.value = vbalProgressBar1.value + 1
SKIP:
        n = n + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    flagchange.Caption = ""
    'vbalProgressBar1.ShowText = False
    'vbalProgressBar1.value = 0
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
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    ReportNameVar = Rptpath & "RPTPURCHREG"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    ''Report.RecordSelectionFormula = "( {ITEMMAST.MANUFACTURER}='" & cmbcompany.Text & "' )"
    Set CRXFormulaFields = Report.FormulaFields
    Dim i As Integer
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
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = " 'Sales Return Register for the Period - ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "PURCHASE REGISTER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    'If Month(Date) > 1 Then
        'CMBMONTH.ListIndex = Month(Date) - 2
    'Else
        'CMBMONTH.ListIndex = 11
    'End If
    OptPetty.Visible = False
    OptAll.Visible = False
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "BILL NO"
    GRDTranx.TextMatrix(0, 2) = "BILL DATE"
    GRDTranx.TextMatrix(0, 3) = "BILL AMT"
    GRDTranx.TextMatrix(0, 4) = "DISC AMT"
    GRDTranx.TextMatrix(0, 5) = "NET AMT"
    GRDTranx.TextMatrix(0, 6) = "CUSTOMER"
    GRDTranx.TextMatrix(0, 7) = ""
    
    GRDTranx.ColWidth(0) = 800
    GRDTranx.ColWidth(1) = 1200
    GRDTranx.ColWidth(2) = 1200
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 4000
    GRDTranx.ColWidth(7) = 0
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 4
    
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
    GRDBILL.ColWidth(4) = 0
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    'Me.Width = 11670
    'Me.Height = 10125
    Me.Left = 1500
    Me.Top = 0
    ACT_FLAG = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            
            FRMEMAIN.Enabled = True
            Frmeperiod.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        FRMEMAIN.Enabled = True
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.rows = 1 Then Exit Sub
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLSUPPLIER.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From RTRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 7) & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
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

Private Sub LBLTRXTOTAL_DblClick()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If OptPetty.Visible = False Then
        OptPetty.Visible = True
        OptAll.Visible = True
    Else
        OptPetty.Visible = False
        OptAll.Visible = False
    End If
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

'Private Sub TMPDELETE_Click()
'    If GRDTranx.rows = 1 Then Exit Sub
'    If MsgBox("Are You Sure You want to Delete BILL NO." & "*** " & GRDTranx.TextMatrix(GRDTranx.Row, 1) & " ****", vbYesNo, "DELETING BILL....") = vbNo Then Exit Sub
'    db.Execute ("DELETE from SALESREG WHERE SALESREG.VCH_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 1) & " AND SALESREG.TRX_TYPE = 'SR'")
'    Call fillSTOCKREG
'End Sub

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
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <>'130000' and ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from CUSTMAST WHERE ACT_CODE <>'130000' and ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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

Private Function fillSTOCKREG()
    Dim rstTRANX As ADODB.Recordset
    Dim TRX_AMOUNT As Double
    Dim i As Long
    
    LBLTRXTOTAL.Caption = ""
    
    On Error GoTo ERRHAND
    TRX_AMOUNT = 0
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")

    Screen.MousePointer = vbHourglass
    TRX_AMOUNT = 0
    
    GRDTranx.rows = 1
    i = 0
    
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    LBLTRXTOTAL.Caption = ""
    GRDTranx.Visible = False
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From SALESREG", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX.EOF
        i = i + 1
        GRDTranx.rows = GRDTranx.rows + 1
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
        MDIMAIN.vbalProgressBar1.Max = rstTRANX.RecordCount
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
    Loop
    
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLTRXTOTAL.Caption = Format(TRX_AMOUNT, "0.00")
    
    MDIMAIN.vbalProgressBar1.ShowText = False
    MDIMAIN.vbalProgressBar1.Value = 0
    GRDTranx.Visible = True
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
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


