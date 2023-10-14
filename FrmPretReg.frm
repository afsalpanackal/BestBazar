VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmPretReg 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE RETURN REGISTER"
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
   Icon            =   "FrmPretReg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
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
      Left            =   495
      TabIndex        =   7
      Top             =   1890
      Visible         =   0   'False
      Width           =   9930
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4005
         Left            =   60
         TabIndex        =   8
         Top             =   540
         Width           =   9795
         _ExtentX        =   17277
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
      BackColor       =   &H00FFFFC0&
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
         BackColor       =   &H00FFFFC0&
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
            Caption         =   "TOTAL BILL AMOUNT"
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
         Cols            =   9
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
         BackColor       =   &H00FFFFC0&
         Caption         =   "Purchase Return Register"
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
            BackColor       =   &H00FFFFC0&
            Height          =   840
            Left            =   7755
            TabIndex        =   25
            Top             =   960
            Width           =   2670
            Begin VB.OptionButton OPTALL 
               BackColor       =   &H00FFFFC0&
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
               BackColor       =   &H00FFFFC0&
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
            BackColor       =   &H00FFFFC0&
            Caption         =   "SUPPLIER"
            Height          =   210
            Left            =   90
            TabIndex        =   19
            Top             =   870
            Width           =   1320
         End
         Begin VB.OptionButton OPTPERIOD 
            BackColor       =   &H00FFFFC0&
            Caption         =   "PERIOD"
            Height          =   210
            Left            =   60
            TabIndex        =   18
            Top             =   420
            Value           =   -1  'True
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
            Format          =   83165185
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
            Format          =   83165185
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
Attribute VB_Name = "FrmPretReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub CMDDISPLAY_Click()
        Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTSALEREG As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim RSTtax As ADODB.Recordset
    Dim N, M As Long
    Dim TAXAMT, EXSALEAMT, TAXSALEAMT, MRPVALUE, DISCAMT As Double
    Dim TAXRATE As Single
    
    db.Execute "delete * From SALESREG"
    
    LBLTRXTOTAL.Caption = "0.00"
    LBLDISCOUNT.Caption = "0.00"
    LBLNET.Caption = "0.00"
    LBLCOST.Caption = "0.00"
    LBLPROFIT.Caption = "0.00"
    lblCommi.Caption = "0.00"
    GRDTranx.Visible = False
    GRDTranx.Rows = 1
    vbalProgressBar1.value = 0
    vbalProgressBar1.ShowText = True
    
    N = 1
    M = 0
    On Error GoTo errHand
    '[BILL_NAME] LIKE '%" & txtCustomerName.Text & "%' AND
    Screen.MousePointer = vbHourglass
    Set rstTRANX = New ADODB.Recordset
    If OPTPERIOD.value = True Then
        rstTRANX.Open "SELECT * From PURCAHSERETURN WHERE [VCH_DATE] <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND [VCH_DATE] >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    Else
        rstTRANX.Open "SELECT * From PURCAHSERETURN WHERE [ACT_CODE] = '" & DataList2.BoundText & "' AND [VCH_DATE] <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND [VCH_DATE] >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " # ORDER BY TRX_TYPE, VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
    End If
        
    If rstTRANX.RecordCount > 0 Then
        vbalProgressBar1.Max = rstTRANX.RecordCount
    Else
        vbalProgressBar1.Max = 100
    End If
    
    Set RSTSALEREG = New ADODB.Recordset
    RSTSALEREG.Open "SELECT * From SALESREG", db, adOpenStatic, adLockOptimistic, adCmdText

    Do Until rstTRANX.EOF
        M = M + 1
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(M, 0) = M
        GRDTranx.TextMatrix(M, 1) = rstTRANX!TRX_TYPE
        GRDTranx.TextMatrix(M, 2) = "CN"
        GRDTranx.TextMatrix(M, 3) = rstTRANX!VCH_NO
        GRDTranx.TextMatrix(M, 4) = rstTRANX!VCH_DATE
        GRDTranx.TextMatrix(M, 5) = Format(Round(rstTRANX!VCH_AMOUNT, 2), "0.00")
'        If rstTRANX!SLSM_CODE = "A" Then
'
'        ElseIf rstTRANX!SLSM_CODE = "P" Then
'            GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(Round((rstTRANX!DISCOUNT * 100 / rstTRANX!VCH_AMOUNT), 2), "0.00"))
'        End If
        GRDTranx.TextMatrix(M, 6) = IIf(IsNull(rstTRANX!DISCOUNT), "", Format(rstTRANX!DISCOUNT, "0.00"))
        GRDTranx.TextMatrix(M, 7) = Format(Round(rstTRANX!NET_AMOUNT, 2), "0.00") 'Format(Round(Val(GRDTranx.TextMatrix(M, 5)) - Val(GRDTranx.TextMatrix(M, 6)), 2), "0.00")
        
        CMDEXIT.Tag = IIf(IsNull(rstTRANX!DISCOUNT), "0", Format(rstTRANX!DISCOUNT, "0.00"))
        'GRDTranx.TextMatrix(M, 7) = Format(Round(Val(GRDTranx.TextMatrix(M, 5)), 2), "0.00")
        GRDTranx.TextMatrix(M, 10) = IIf(IsNull(rstTRANX!ACT_NAME), "", rstTRANX!ACT_NAME)
        GRDTranx.TextMatrix(M, 11) = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME) & IIf(IsNull(rstTRANX!BILL_ADDRESS), "", ", " & rstTRANX!BILL_ADDRESS)

        CMDDISPLAY.Tag = ""
        frmemain.Tag = ""
        FRMEBILL.Tag = ""
        
        'If rstTRANX!TRX_TYPE <> "SI" Then GoTo SKIP
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select DISTINCT SALES_TAX From TRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            EXSALEAMT = 0
            TAXSALEAMT = 0
            TAXAMT = 0
            MRPVALUE = 0
            DISCAMT = 0
            TAXRATE = RSTTRXFILE!SALES_TAX
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & RSTTRXFILE!SALES_TAX & "", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                If RSTTRXFILE!SALES_TAX > 0 And RSTtax!CHECK_FLAG = "V" Then
                    TAXSALEAMT = TAXSALEAMT + IIf(IsNull(RSTtax!TRX_TOTAL), 0, RSTtax!TRX_TOTAL)
                    'TAXAMT = TAXAMT + Round((RSTtax!PTR * RSTtax!SALES_TAX / 100) * RSTtax!QTY, 2)
                    TAXAMT = TAXAMT + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                Else
                    If RSTtax!SALE_1_FLAG = "1" Then
                        TAXAMT = TAXAMT + Round((RSTtax!SALES_PRICE - RSTtax!PTR) * RSTtax!QTY, 2)
                        MRPVALUE = Round(MRPVALUE + (100 * RSTtax!MRP / 105) * RSTtax!QTY, 2)
                    End If
                    EXSALEAMT = EXSALEAMT + RSTtax!TRX_TOTAL
                End If
                DISCAMT = Round(DISCAMT + IIf(IsNull(RSTtax!LINE_DISC), 0, RSTtax!TRX_TOTAL * RSTtax!LINE_DISC / 100), 2)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            RSTSALEREG.AddNew
            TAXSALEAMT = TAXSALEAMT - TAXAMT
            RSTSALEREG!VCH_NO = rstTRANX!VCH_NO 'N
            RSTSALEREG!TRX_TYPE = rstTRANX!TRX_TYPE
            RSTSALEREG!VCH_DATE = rstTRANX!VCH_DATE
            RSTSALEREG!DISCOUNT = DISCAMT
            RSTSALEREG!VCH_AMOUNT = Val(GRDTranx.TextMatrix(M, 7))
            RSTSALEREG!PAYAMOUNT = Val(GRDTranx.TextMatrix(M, 9))
            RSTSALEREG!ACT_NAME = IIf(IsNull(rstTRANX!BILL_NAME), "", rstTRANX!BILL_NAME)
            RSTSALEREG!ACT_CODE = IIf(IsNull(rstTRANX!ACT_CODE), "", rstTRANX!ACT_CODE)
            
            Dim RSTACTCODE As ADODB.Recordset
            Set RSTACTCODE = New ADODB.Recordset
            RSTACTCODE.Open "SELECT [KGST] FROM CUSTMAST WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTACTCODE.EOF And RSTACTCODE.BOF) Then
                RSTSALEREG!TIN_NO = RSTACTCODE!KGST
            End If
            RSTACTCODE.Close
            Set RSTACTCODE = Nothing
            RSTSALEREG!EXMPSALES_AMT = EXSALEAMT
            RSTSALEREG!TAXSALES_AMT = TAXSALEAMT
            RSTSALEREG!TAXAMOUNT = TAXAMT
            RSTSALEREG!TAXRATE = TAXRATE
            CMDDISPLAY.Tag = Val(CMDDISPLAY.Tag) + EXSALEAMT
            frmemain.Tag = Val(frmemain.Tag) + TAXSALEAMT
            FRMEBILL.Tag = Val(FRMEBILL.Tag) + TAXAMT
            RSTSALEREG.Update
            
            RSTTRXFILE.MoveNext
        Loop
    
        GRDTranx.TextMatrix(M, 12) = Format(Val(CMDDISPLAY.Tag), "0.00")
        GRDTranx.TextMatrix(M, 13) = Format(Val(frmemain.Tag), "0.00")
        GRDTranx.TextMatrix(M, 14) = Format(Val(FRMEBILL.Tag), "0.00")
        
        LBLTRXTOTAL.Caption = Format(Val(LBLTRXTOTAL.Caption) + rstTRANX!VCH_AMOUNT, "0.00")
        LBLDISCOUNT.Caption = Format(Val(LBLDISCOUNT.Caption) + Val(GRDTranx.TextMatrix(M, 6)), "0.00")
        If frmLogin.rs!Level <> "0" Then
            lblCommi.Caption = "xxx"
            LBLCOST.Caption = "xxx"
            LBLPROFIT.Caption = "xxx"
        Else
            lblCommi.Caption = Format(Val(lblCommi.Caption) + Val(GRDTranx.TextMatrix(M, 8)), "0.00")
            LBLCOST.Caption = Format(Val(LBLCOST.Caption) + Val(GRDTranx.TextMatrix(M, 9)), "0.00")
            'LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblcommi.Caption)), "0.00")
        End If
        vbalProgressBar1.value = vbalProgressBar1.value + 1
SKIP:
        N = N + 1
        rstTRANX.MoveNext
    Loop

    rstTRANX.Close
    Set rstTRANX = Nothing
    RSTSALEREG.Close
    Set RSTSALEREG = Nothing
    
    LBLNET.Caption = Format(Val(LBLTRXTOTAL.Caption) - Val(LBLDISCOUNT.Caption), "0.00")
    If frmLogin.rs!Level = "0" Then
        LBLPROFIT.Caption = Format(Val(LBLNET.Caption) - (Val(LBLCOST.Caption) + Val(lblCommi.Caption)), "0.00")
    End If
        
    flagchange.Caption = ""
    vbalProgressBar1.ShowText = False
    vbalProgressBar1.value = 0
    GRDTranx.Visible = True
    CMDPRINTREGISTER.Enabled = True
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Sub
    
errHand:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDPRINTREGISTER_Click()
    rptPRINT.ReportFileName = "D:\EzBiz\RPTSALESREG.RPT"
    rptPRINT.Formulas(0) = "PERIOD = '" & DTFROM.value & " " & " TO " & " " & DTTO.value & "'"
    rptPRINT.Action = 1
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "SL"
    GRDTranx.TextMatrix(0, 1) = "P/R NO"
    GRDTranx.TextMatrix(0, 2) = "TYPE"
    GRDTranx.TextMatrix(0, 3) = "BILL DATE"
    GRDTranx.TextMatrix(0, 4) = "AMOUNT"
    GRDTranx.TextMatrix(0, 5) = "RETURNED TO"
    GRDTranx.TextMatrix(0, 6) = "STATUS"
    GRDTranx.TextMatrix(0, 7) = "INVOICE NO."
    GRDTranx.TextMatrix(0, 8) = "INVOICE DATE"
    
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 0
    GRDTranx.ColWidth(3) = 1350
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 2500
    GRDTranx.ColWidth(6) = 1100
    GRDTranx.ColWidth(7) = 1100
    GRDTranx.ColWidth(8) = 1100
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 6
    GRDTranx.ColAlignment(5) = 1
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(7) = 6
    GRDTranx.ColAlignment(8) = 6
    
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "MRP"
    GRDBILL.TextMatrix(0, 3) = "Rate"
    GRDBILL.TextMatrix(0, 4) = "Qty"
    GRDBILL.TextMatrix(0, 5) = "Amount"
    GRDBILL.TextMatrix(0, 6) = "Serial No"
    GRDBILL.TextMatrix(0, 7) = ""
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2500
    GRDBILL.ColWidth(2) = 900
    GRDBILL.ColWidth(3) = 900
    GRDBILL.ColWidth(4) = 900
    GRDBILL.ColWidth(5) = 1100
    GRDBILL.ColWidth(6) = 1100
    GRDBILL.ColWidth(7) = 0
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 6
    GRDBILL.ColAlignment(4) = 6
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    GRDBILL.ColAlignment(7) = 1
    
    CLOSEALL = 1
    ACT_FLAG = True
    Month (Date) - 2
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    DTTO.value = Format(Date, "DD/MM/YYYY")
    Me.Width = 10845
    Me.Height = 11025
    Me.Left = 1500
    Me.Top = 0
    txtpassword = "YEAR " & Year(Date)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            
            frmemain.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.Rows = 1 Then Exit Sub
            LBLBILLNO.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 1)
            LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
             
            GRDBILL.Rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From PURCAHSERETURN WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & Trim(GRDTranx.TextMatrix(GRDTranx.Row, 2)) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!MRP, "0.00")
                GRDBILL.TextMatrix(i, 3) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 5) = Format(RSTTRXFILE!SALES_PRICE * RSTTRXFILE!QTY, "0.00")
                GRDBILL.TextMatrix(i, 6) = RSTTRXFILE!REF_NO
                GRDBILL.TextMatrix(i, 7) = "" 'IIf(IsNull(RSTTRXFILE!EXP_DATE), "", RSTTRXFILE!EXP_DATE)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            frmemain.Enabled = False
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
    OPTCUSTOMER.value = True
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo errHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly
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
errHand:
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
        Case Asc("'"), Asc("["), Asc("]"), Asc("["), Asc("]")
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

