VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTaxdetails 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tax Details"
   ClientHeight    =   8370
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16350
   Icon            =   "frmtaxdetails.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8370
   ScaleWidth      =   16350
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8400
      Left            =   0
      TabIndex        =   0
      Top             =   -90
      Width           =   16350
      Begin VB.CommandButton CmDDisplay 
         Caption         =   "&Display"
         Height          =   420
         Left            =   5805
         TabIndex        =   8
         Top             =   165
         Width           =   1065
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         Height          =   420
         Left            =   6930
         TabIndex        =   7
         Top             =   165
         Width           =   1065
      End
      Begin VB.OptionButton OPTPERIOD 
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
         Height          =   210
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Value           =   -1  'True
         Width           =   1620
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   4620
         Left            =   45
         TabIndex        =   1
         Top             =   660
         Width           =   7980
         _ExtentX        =   14076
         _ExtentY        =   8149
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx2 
         Height          =   2325
         Left            =   8055
         TabIndex        =   2
         Top             =   660
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   4101
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   1680
         TabIndex        =   4
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   145293313
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   3570
         TabIndex        =   5
         Top             =   180
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   145293313
         CurrentDate     =   40498
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx3 
         Height          =   2265
         Left            =   8055
         TabIndex        =   15
         Top             =   3015
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   3995
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx4 
         Height          =   2355
         Left            =   8055
         TabIndex        =   16
         Top             =   5310
         Width           =   8250
         _ExtentX        =   14552
         _ExtentY        =   4154
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx5 
         Height          =   2355
         Left            =   60
         TabIndex        =   17
         Top             =   5310
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   4154
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
         AllowUserResizing=   3
         Appearance      =   0
         GridLineWidth   =   2
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLDIFF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   6420
         TabIndex        =   14
         Top             =   7725
         Width           =   1605
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Diff. Tax"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Index           =   0
         Left            =   5460
         TabIndex        =   13
         Top             =   7740
         Width           =   930
      End
      Begin VB.Label LBLINTAX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   1095
         TabIndex        =   12
         Top             =   7725
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Input Tax"
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
         Index           =   4
         Left            =   90
         TabIndex        =   11
         Top             =   7740
         Width           =   1005
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Output Tax"
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
         Index           =   3
         Left            =   2595
         TabIndex        =   10
         Top             =   7740
         Width           =   1110
      End
      Begin VB.Label LBLOPTAX 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   3690
         TabIndex        =   9
         Top             =   7725
         Width           =   1605
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
         Left            =   3285
         TabIndex        =   6
         Top             =   240
         Width           =   285
      End
   End
End
Attribute VB_Name = "FrmTaxdetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CmDDisplay_Click()
    
    Dim KFC As Double
    Dim CESSPER As Double
    Dim CESSAMT As Double
    Dim TAX_AMT As Double
    Dim TAXABLE_AMT As Double
    Dim NET_AMT As Double
    
    Dim T_CESSPER As Double
    Dim T_CESSAMT As Double
    Dim T_TAX_AMT As Double
    Dim T_TAXABLE_AMT As Double
    Dim T_NET_AMT As Double
    
    Dim F_CESSPER As Double
    Dim F_CESSAMT As Double
    Dim F_TAX_AMT As Double
    Dim F_TAXABLE_AMT As Double
    Dim F_NET_AMT As Double
    
    LBLINTAX.Caption = ""
    LBLOPTAX.Caption = ""
    LBLDIFF.Caption = ""
    
    CESSPER = 0
    CESSAMT = 0
    KFC = 0
    TAX_AMT = 0
    TAXABLE_AMT = 0
    NET_AMT = 0
    
    F_CESSPER = 0
    F_CESSAMT = 0
    F_TAX_AMT = 0
    F_TAXABLE_AMT = 0
    F_NET_AMT = 0
    
    Dim RSTtax As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim rstTRANX2 As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer
    
    Screen.MousePointer = vbHourglass
    
    i = 1
    
    GRDTranx.FixedRows = 0
    GRDTranx.Rows = 1
    
    
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.FixedRows = 1
    GRDTranx.TextMatrix(i, 0) = "Sales"
    i = i + 1
            
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.FixedRows = 1
    GRDTranx.TextMatrix(i, 0) = "B2C Sales"
    i = i + 1
    
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
    
    On Error GoTo Errhand
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'B2C
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='HI'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTTRXFILE!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                    KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!KFC_TAX), 0, RSTtax!KFC_TAX / 100)) * RSTtax!QTY
                    TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    NET_AMT = NET_AMT + TAXABLE_AMT + TAX_AMT
                    
                    CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                    
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
        
        F_CESSPER = F_CESSPER + CESSPER
        F_CESSAMT = F_CESSAMT + CESSAMT
        F_TAX_AMT = F_TAX_AMT + TAX_AMT
        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
        F_NET_AMT = F_NET_AMT + NET_AMT
                
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.TextMatrix(i, 0) = "TOTAL"
    GRDTranx.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    i = i + 1
                
                    
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
                
    '=======B2B
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.FixedRows = 1
    GRDTranx.TextMatrix(i, 0) = "B2B Sales"
    i = i + 1
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'B2B
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='GI'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTTRXFILE!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                    KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!KFC_TAX), 0, RSTtax!KFC_TAX / 100)) * RSTtax!QTY
                    TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    NET_AMT = NET_AMT + TAXABLE_AMT + TAX_AMT
                    
                    CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                    
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
    
                        
        F_CESSPER = F_CESSPER + CESSPER
        F_CESSAMT = F_CESSAMT + CESSAMT
        F_TAX_AMT = F_TAX_AMT + TAX_AMT
        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
        F_NET_AMT = F_NET_AMT + NET_AMT
        
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.TextMatrix(i, 0) = "TOTAL"
    GRDTranx.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    
    i = i + 1
        
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
    
    '=======SERVICE BILLS
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.FixedRows = 1
    GRDTranx.TextMatrix(i, 0) = "SERVICE BILLS"
    i = i + 1
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'SERVICE BILLS
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRXMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SV'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From TRXMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTTRXFILE!SLSM_CODE
                        Case "P"
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                        Case Else
                            GRDTranx.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx.Tag) * Val(RSTtax!QTY)
                    KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!KFC_TAX), 0, RSTtax!KFC_TAX / 100)) * RSTtax!QTY
                    TAX_AMT = TAX_AMT + (Val(GRDTranx.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    NET_AMT = NET_AMT + TAXABLE_AMT + TAX_AMT
                    
                    CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                    
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx.Rows = GRDTranx.Rows + 1
            GRDTranx.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
    
        F_CESSPER = F_CESSPER + CESSPER
        F_CESSAMT = F_CESSAMT + CESSAMT
        F_TAX_AMT = F_TAX_AMT + TAX_AMT
        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
        F_NET_AMT = F_NET_AMT + NET_AMT
        
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.TextMatrix(i, 0) = "TOTAL"
    GRDTranx.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    
    i = i + 1
    
    GRDTranx.Rows = GRDTranx.Rows + 1
    GRDTranx.TextMatrix(i, 0) = "G. TOTAL"
    GRDTranx.TextMatrix(i, 1) = Format(Round(F_TAXABLE_AMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 2) = Format(Round(F_TAX_AMT, 2), "0.00")
    
    GRDTranx.TextMatrix(i, 3) = Format(Round(F_CESSPER, 2), "0.00")
    GRDTranx.TextMatrix(i, 4) = Format(Round(F_CESSAMT, 2), "0.00")
    GRDTranx.TextMatrix(i, 5) = Format(Round(F_TAXABLE_AMT + F_TAX_AMT + F_CESSPER + F_CESSAMT, 2), "0.00")
    
    LBLOPTAX.Caption = Format(F_TAX_AMT, "0.00")
    
    i = 1
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
        
    GRDTranx2.Rows = 2
    GRDTranx2.FixedRows = 1
    GRDTranx2.TextMatrix(i, 0) = "PURCHASE"
    i = i + 1
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='PI' ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'PURCHASE
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From TRANSMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='PI'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(CATEGORY) OR CATEGORY <> 'SERVICE CHARGE')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From TRANSMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTtax!DISC_FLAG
                        Case "P"
                            GRDTranx2.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!TR_DISC / 100) '- ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx2.Tag = RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY)) - ((RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY))) * RSTtax!TR_DISC / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx2.Tag) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
                    TAX_AMT = TAX_AMT + (Val(GRDTranx2.Tag) * RSTtax!SALES_TAX / 100) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
    
                    If RSTtax!DISC_FLAG = "P" Then
                        CESSPER = CESSPER + (Val(GRDTranx2.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    Else
                        CESSPER = CESSPER + (Val(GRDTranx2.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    End If
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx2.Rows = GRDTranx2.Rows + 1
            GRDTranx2.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx2.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx2.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx2.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx2.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx2.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
        
'        F_CESSPER = F_CESSPER + CESSPER
'        F_CESSAMT = F_CESSAMT + CESSAMT
'        F_TAX_AMT = F_TAX_AMT + TAX_AMT
'        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
'        F_NET_AMT = F_NET_AMT + NET_AMT
                
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx2.Rows = GRDTranx2.Rows + 1
    GRDTranx2.TextMatrix(i, 0) = "TOTAL"
    GRDTranx2.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx2.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx2.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx2.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx2.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    
    LBLINTAX.Caption = Format(T_TAX_AMT, "0.00")
    
    'i = i + 1
    
    i = 1
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
        
    GRDTranx3.Rows = 2
    GRDTranx3.FixedRows = 1
    GRDTranx3.TextMatrix(i, 0) = "SALES RETURN"
    i = i + 1
    
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR' ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'SALES RETURN
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From RETURNMAST WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='SR'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From RTRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(CATEGORY) OR CATEGORY <> 'SERVICE CHARGE')", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From RETURNMAST WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTtax!DISC_FLAG
                        Case "P"
                            GRDTranx3.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) ' - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * RSTtax!TR_DISC / 100) '- ((RSTtax!PTR - (RSTtax!PTR * RSTtax!P_DISC) / 100) * IIf(IsNull(rstTRANX!DISC_PERS), 0, rstTRANX!DISC_PERS) / 100)
                        Case Else
                            GRDTranx3.Tag = RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY)) '- ((RSTtax!PTR - (RSTtax!P_DISC / Val(RSTtax!QTY))) * RSTtax!TR_DISC / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx3.Tag) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
                    TAX_AMT = TAX_AMT + (Val(GRDTranx3.Tag) * RSTtax!SALES_TAX / 100) * (Val(RSTtax!QTY) - IIf(IsNull(RSTtax!SCHEME), 0, Val(RSTtax!SCHEME)))
    
                    If RSTtax!DISC_FLAG = "P" Then
                        CESSPER = CESSPER + (Val(GRDTranx3.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    Else
                        CESSPER = CESSPER + (Val(GRDTranx3.Tag) * RSTtax!QTY) * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    End If
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx3.Rows = GRDTranx3.Rows + 1
            GRDTranx3.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx3.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx3.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx3.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx3.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx3.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
        
'        F_CESSPER = F_CESSPER + CESSPER
'        F_CESSAMT = F_CESSAMT + CESSAMT
'        F_TAX_AMT = F_TAX_AMT + TAX_AMT
'        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
'        F_NET_AMT = F_NET_AMT + NET_AMT
                
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx3.Rows = GRDTranx3.Rows + 1
    GRDTranx3.TextMatrix(i, 0) = "TOTAL"
    GRDTranx3.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx3.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx3.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx3.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx3.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    LBLINTAX.Caption = Val(LBLINTAX.Caption) + T_TAX_AMT
    
    i = 1
    
    GRDTranx4.FixedRows = 0
    GRDTranx4.Rows = 1
    
    
    GRDTranx4.Rows = GRDTranx4.Rows + 1
    GRDTranx4.FixedRows = 1
    GRDTranx4.TextMatrix(i, 0) = "PURCHASE RETURN"
    i = i + 1
            
   
    
    T_CESSPER = 0
    T_CESSAMT = 0
    T_TAX_AMT = 0
    T_TAXABLE_AMT = 0
    T_NET_AMT = 0
    
    
    Set rstTRANX2 = New ADODB.Recordset
    rstTRANX2.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='PR' ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    Do Until rstTRANX2.EOF
    
        'PURCHASE RETURN
        
        CESSPER = 0
        CESSAMT = 0
        KFC = 0
        TAX_AMT = 0
        TAXABLE_AMT = 0
        NET_AMT = 0
        
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT * From PURCAHSERETURN WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_TYPE='PR'  ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            
            
            Set RSTtax = New ADODB.Recordset
            RSTtax.Open "Select * From TRXFILE WHERE VCH_DATE <= '" & Format(DTTO.value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' AND SALES_TAX = " & rstTRANX2!SALES_TAX & " AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTtax.EOF
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT * From PURCAHSERETURN WHERE VCH_NO =" & RSTtax!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ORDER BY VCH_NO", db, adOpenStatic, adLockReadOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    Select Case RSTTRXFILE!SLSM_CODE
                        Case "P"
                            GRDTranx4.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) - ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTTRXFILE!DISC_PERS), 0, RSTTRXFILE!DISC_PERS) / 100)
                        Case Else
                            GRDTranx4.Tag = (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100)
                    End Select
                    TAXABLE_AMT = TAXABLE_AMT + Val(GRDTranx4.Tag) * Val(RSTtax!QTY)
                    KFC = KFC + ((RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * IIf(IsNull(RSTtax!KFC_TAX), 0, RSTtax!KFC_TAX / 100)) * RSTtax!QTY
                    TAX_AMT = TAX_AMT + (Val(GRDTranx4.Tag) * RSTtax!SALES_TAX / 100) * RSTtax!QTY
                    NET_AMT = NET_AMT + TAXABLE_AMT + TAX_AMT
                    
                    CESSPER = CESSPER + (RSTtax!PTR - (RSTtax!PTR * RSTtax!LINE_DISC) / 100) * RSTtax!QTY * IIf(IsNull(RSTtax!CESS_PER), 0, RSTtax!CESS_PER / 100)
                    CESSAMT = CESSAMT + IIf(IsNull(RSTtax!CESS_AMT), 0, RSTtax!CESS_AMT) * RSTtax!QTY
                    
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                'GRDPranx.TextMatrix(M, N) = Val(GRDPranx.TextMatrix(M, N)) + Val(GRDPranx.Tag) * Val(RSTtax!QTY)
                RSTtax.MoveNext
            Loop
            RSTtax.Close
            Set RSTtax = Nothing
            
            GRDTranx4.Rows = GRDTranx4.Rows + 1
            GRDTranx4.TextMatrix(i, 0) = rstTRANX2!SALES_TAX & "%"
            GRDTranx4.TextMatrix(i, 1) = Format(Round(TAXABLE_AMT, 2), "0.00")
            GRDTranx4.TextMatrix(i, 2) = Format(Round(TAX_AMT, 2), "0.00")
            
            GRDTranx4.TextMatrix(i, 3) = Format(Round(CESSPER, 2), "0.00")
            GRDTranx4.TextMatrix(i, 4) = Format(Round(CESSAMT, 2), "0.00")
            GRDTranx4.TextMatrix(i, 5) = Format(Round(TAXABLE_AMT + TAX_AMT + CESSPER + CESSAMT, 2), "0.00")
            
            i = i + 1
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
            
        T_CESSPER = T_CESSPER + CESSPER
        T_CESSAMT = T_CESSAMT + CESSAMT
        T_TAX_AMT = T_TAX_AMT + TAX_AMT
        T_TAXABLE_AMT = T_TAXABLE_AMT + TAXABLE_AMT
        T_NET_AMT = T_NET_AMT + NET_AMT
        
        F_CESSPER = F_CESSPER + CESSPER
        F_CESSAMT = F_CESSAMT + CESSAMT
        F_TAX_AMT = F_TAX_AMT + TAX_AMT
        F_TAXABLE_AMT = F_TAXABLE_AMT + TAXABLE_AMT
        F_NET_AMT = F_NET_AMT + NET_AMT
                
        rstTRANX2.MoveNext
    Loop
    rstTRANX2.Close
    Set rstTRANX2 = Nothing
    
    GRDTranx4.Rows = GRDTranx4.Rows + 1
    GRDTranx4.TextMatrix(i, 0) = "TOTAL"
    GRDTranx4.TextMatrix(i, 1) = Format(Round(T_TAXABLE_AMT, 2), "0.00")
    GRDTranx4.TextMatrix(i, 2) = Format(Round(T_TAX_AMT, 2), "0.00")
    
    GRDTranx4.TextMatrix(i, 3) = Format(Round(T_CESSPER, 2), "0.00")
    GRDTranx4.TextMatrix(i, 4) = Format(Round(T_CESSAMT, 2), "0.00")
    GRDTranx4.TextMatrix(i, 5) = Format(Round(T_TAXABLE_AMT + T_TAX_AMT + T_CESSPER + T_CESSAMT, 2), "0.00")
    i = i + 1
                
                    
    
    LBLOPTAX.Caption = Val(LBLOPTAX.Caption) + T_TAX_AMT
    
    LBLINTAX.Caption = Format(Round(LBLINTAX.Caption, 2), "0.00")
    LBLOPTAX.Caption = Format(Round(LBLOPTAX.Caption, 2), "0.00")
    LBLDIFF.Caption = Format(Round(Val(LBLINTAX.Caption) - Val(LBLOPTAX.Caption), 2), "0.00")
    Screen.MousePointer = vbNormal
    Exit Sub
Errhand:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
    
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    GRDTranx.TextMatrix(0, 0) = "Tax %"
    GRDTranx.TextMatrix(0, 1) = "Taxable Value"
    GRDTranx.TextMatrix(0, 2) = "Tax Amt"
    GRDTranx.TextMatrix(0, 3) = "Cess Amt"
    GRDTranx.TextMatrix(0, 4) = "Addl Cess"
    GRDTranx.TextMatrix(0, 5) = "Net Amount"
    
    GRDTranx.ColWidth(0) = 900
    GRDTranx.ColWidth(1) = 1600
    GRDTranx.ColWidth(2) = 1400
    GRDTranx.ColWidth(3) = 1000
    GRDTranx.ColWidth(4) = 1000
    GRDTranx.ColWidth(5) = 1600
    
    
    GRDTranx.ColAlignment(0) = 4
    GRDTranx.ColAlignment(1) = 1
    GRDTranx.ColAlignment(2) = 1
    GRDTranx.ColAlignment(3) = 1
    GRDTranx.ColAlignment(4) = 1
    GRDTranx.ColAlignment(5) = 1
    
    
    GRDTranx.FixedRows = 0
    GRDTranx.Rows = 1
    
    
    GRDTranx2.TextMatrix(0, 0) = "Tax %"
    GRDTranx2.TextMatrix(0, 1) = "Taxable Value"
    GRDTranx2.TextMatrix(0, 2) = "Tax Amt"
    GRDTranx2.TextMatrix(0, 3) = "Cess Amt"
    GRDTranx2.TextMatrix(0, 4) = "Addl Cess"
    GRDTranx2.TextMatrix(0, 5) = "Net Amount"
    
    GRDTranx2.ColWidth(0) = 900
    GRDTranx2.ColWidth(1) = 1600
    GRDTranx2.ColWidth(2) = 1400
    GRDTranx2.ColWidth(3) = 1000
    GRDTranx2.ColWidth(4) = 1000
    GRDTranx2.ColWidth(5) = 1600
    
    
    GRDTranx2.ColAlignment(0) = 4
    GRDTranx2.ColAlignment(1) = 1
    GRDTranx2.ColAlignment(2) = 1
    GRDTranx2.ColAlignment(3) = 1
    GRDTranx2.ColAlignment(4) = 1
    GRDTranx2.ColAlignment(5) = 1
    
    
    GRDTranx2.FixedRows = 0
    GRDTranx2.Rows = 1
    
    
    GRDTranx3.TextMatrix(0, 0) = "Tax %"
    GRDTranx3.TextMatrix(0, 1) = "Taxable Value"
    GRDTranx3.TextMatrix(0, 2) = "Tax Amt"
    GRDTranx3.TextMatrix(0, 3) = "Cess Amt"
    GRDTranx3.TextMatrix(0, 4) = "Addl Cess"
    GRDTranx3.TextMatrix(0, 5) = "Net Amount"
    
    GRDTranx3.ColWidth(0) = 900
    GRDTranx3.ColWidth(1) = 1600
    GRDTranx3.ColWidth(2) = 1400
    GRDTranx3.ColWidth(3) = 1000
    GRDTranx3.ColWidth(4) = 1000
    GRDTranx3.ColWidth(5) = 1600
    
    
    GRDTranx3.ColAlignment(0) = 4
    GRDTranx3.ColAlignment(1) = 1
    GRDTranx3.ColAlignment(2) = 1
    GRDTranx3.ColAlignment(3) = 1
    GRDTranx3.ColAlignment(4) = 1
    GRDTranx3.ColAlignment(5) = 1
    
    
    GRDTranx3.FixedRows = 0
    GRDTranx3.Rows = 1
    
    
    GRDTranx2.FixedRows = 0
    GRDTranx2.Rows = 1
    
    
    GRDTranx4.TextMatrix(0, 0) = "Tax %"
    GRDTranx4.TextMatrix(0, 1) = "Taxable Value"
    GRDTranx4.TextMatrix(0, 2) = "Tax Amt"
    GRDTranx4.TextMatrix(0, 3) = "Cess Amt"
    GRDTranx4.TextMatrix(0, 4) = "Addl Cess"
    GRDTranx4.TextMatrix(0, 5) = "Net Amount"
    
    GRDTranx4.ColWidth(0) = 900
    GRDTranx4.ColWidth(1) = 1600
    GRDTranx4.ColWidth(2) = 1400
    GRDTranx4.ColWidth(3) = 1000
    GRDTranx4.ColWidth(4) = 1000
    GRDTranx4.ColWidth(5) = 1600
    
    
    GRDTranx4.ColAlignment(0) = 4
    GRDTranx4.ColAlignment(1) = 1
    GRDTranx4.ColAlignment(2) = 1
    GRDTranx4.ColAlignment(3) = 1
    GRDTranx4.ColAlignment(4) = 1
    GRDTranx4.ColAlignment(5) = 1
    
    
    GRDTranx4.FixedRows = 0
    GRDTranx4.Rows = 1
    DTFROM.value = "01" & "/" & Month(Date) & "/" & Year(Date)
    DTTO.value = Format(Date, "DD/MM/YYYY")
    
    Me.Left = 1000
    Me.Top = 0
    
End Sub

Private Function SALES_RETURN()

    
    Screen.MousePointer = vbNormal
    Exit Function
    Screen.MousePointer = vbNormal
Errhand:
    MsgBox Err.Description
    
End Function


