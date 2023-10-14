VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmeasybill 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E3FBF5&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Easy Bill"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18900
   DrawMode        =   3  'Not Merge Pen
   Icon            =   "FrmZeroStk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   18900
   Begin VB.Frame Frame4 
      Caption         =   "Total Purchase"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6420
      TabIndex        =   30
      Top             =   7560
      Width           =   3135
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "For the month"
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
         Height          =   375
         Index           =   7
         Left            =   45
         TabIndex        =   32
         Top             =   315
         Width           =   1470
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblpurcahsemonth 
         Alignment       =   2  'Center
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
         Height          =   420
         Left            =   1515
         TabIndex        =   31
         Top             =   225
         Width           =   1575
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Total Sales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   15
      TabIndex        =   25
      Top             =   7560
      Width           =   6390
      Begin VB.Label lblsalesmonth 
         Alignment       =   2  'Center
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
         Height          =   420
         Left            =   4560
         TabIndex        =   29
         Top             =   225
         Width           =   1725
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "For the month"
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
         Height          =   375
         Index           =   5
         Left            =   3105
         TabIndex        =   28
         Top             =   315
         Width           =   1425
         WordWrap        =   -1  'True
      End
      Begin VB.Label LBLsalesday 
         Alignment       =   2  'Center
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
         Height          =   420
         Left            =   1515
         TabIndex        =   27
         Top             =   225
         Width           =   1575
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "For the date"
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
         Height          =   375
         Index           =   50
         Left            =   45
         TabIndex        =   26
         Top             =   315
         Width           =   1470
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox tXTITEM 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1200
      TabIndex        =   0
      Top             =   75
      Width           =   4080
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   7020
      Left            =   -15
      TabIndex        =   21
      Top             =   525
      Width           =   12345
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
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   6990
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   12345
         _ExtentX        =   21775
         _ExtentY        =   12330
         _Version        =   393216
         Rows            =   1
         Cols            =   17
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   7815
      TabIndex        =   16
      Top             =   8310
      Width           =   1170
   End
   Begin VB.CommandButton CmdExit 
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
      Height          =   450
      Left            =   10260
      TabIndex        =   15
      Top             =   8310
      Width           =   1200
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "Print Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   9015
      TabIndex        =   14
      Top             =   8310
      Width           =   1200
   End
   Begin MSMask.MaskEdBox TXTINVDATE 
      Height          =   300
      Left            =   10800
      TabIndex        =   3
      Top             =   60
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.CommandButton CmdMake 
      Caption         =   "Make Bill"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   6660
      TabIndex        =   2
      Top             =   8310
      Width           =   1095
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   5460
      TabIndex        =   1
      Top             =   8310
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F1EBDC&
      Caption         =   "Billing Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1470
      Left            =   12315
      TabIndex        =   9
      Top             =   30
      Width           =   3840
      Begin VB.TextBox TxtBillName 
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
         Left            =   45
         MaxLength       =   100
         TabIndex        =   10
         Top             =   225
         Width           =   3750
      End
      Begin MSForms.TextBox TxtBillAddress 
         Height          =   840
         Left            =   45
         TabIndex        =   11
         Top             =   570
         Width           =   3750
         VariousPropertyBits=   -1400879077
         MaxLength       =   150
         BorderStyle     =   1
         Size            =   "6615;1482"
         SpecialEffect   =   0
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame5"
      Height          =   4545
      Left            =   12345
      TabIndex        =   33
      Top             =   1500
      Width           =   6555
      Begin VB.TextBox txtsales 
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
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GrdSales 
         Height          =   4485
         Left            =   -15
         TabIndex        =   34
         Top             =   15
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   7911
         _Version        =   393216
         Rows            =   1
         Cols            =   16
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   400
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         AllowBigSelection=   0   'False
         FocusRect       =   2
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
   Begin MSFlexGridLib.MSFlexGrid GRDTranx 
      Height          =   1350
      Left            =   12345
      TabIndex        =   36
      Top             =   6060
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2381
      _Version        =   393216
      Cols            =   19
      FixedCols       =   0
      RowHeightMin    =   350
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
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
   Begin MSFlexGridLib.MSFlexGrid GRDTranx2 
      Height          =   1350
      Left            =   12345
      TabIndex        =   37
      Top             =   7440
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   2381
      _Version        =   393216
      Cols            =   19
      FixedCols       =   0
      RowHeightMin    =   350
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
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
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   4
      Left            =   60
      TabIndex        =   24
      Top             =   120
      Width           =   1260
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Cost"
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
      Height          =   375
      Index           =   3
      Left            =   16170
      TabIndex        =   20
      Top             =   1065
      Width           =   1440
   End
   Begin VB.Label LBLTOTALCOST 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   17295
      TabIndex        =   19
      Top             =   990
      Width           =   1545
   End
   Begin VB.Label LBLTOTAL 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   17295
      TabIndex        =   18
      Top             =   45
      Width           =   1545
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Gross Amt"
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
      Height          =   375
      Index           =   2
      Left            =   16170
      TabIndex        =   17
      Top             =   120
      Width           =   1485
   End
   Begin VB.Label lblselAMT 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   9675
      TabIndex        =   13
      Top             =   7815
      Width           =   1800
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Amount"
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
      Height          =   375
      Index           =   1
      Left            =   9660
      TabIndex        =   12
      Top             =   7560
      Width           =   1785
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Amt"
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
      Height          =   375
      Index           =   23
      Left            =   16170
      TabIndex        =   8
      Top             =   630
      Width           =   1440
   End
   Begin VB.Label lblnetamount 
      Alignment       =   2  'Center
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
      Height          =   405
      Left            =   17295
      TabIndex        =   7
      Top             =   525
      Width           =   1545
   End
   Begin VB.Label LBLBILLNO 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   9165
      TabIndex        =   6
      Top             =   30
      Width           =   1050
   End
   Begin VB.Label Label1 
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
      Height          =   255
      Index           =   0
      Left            =   8340
      TabIndex        =   5
      Top             =   45
      Width           =   780
   End
   Begin VB.Label INVDATE 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   8
      Left            =   10260
      TabIndex        =   4
      Top             =   60
      Width           =   630
   End
End
Attribute VB_Name = "frmeasybill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SAVE_BILL As Boolean

Private Sub CmDDisplay_Click()
    Dim rststock, RSTRTRXFILE, RSTSUPPLIER As ADODB.Recordset
    Dim i As Long
    
    'PHY_FLAG = True
    
    
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.rows = 1
    i = 1
    
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    Dim RSTITEMMAST As ADODB.Recordset
    Set rststock = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenForwardOnly
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM RTRXFILE where bal_qty >0 and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y') ORDER BY trx_type, trx_year, vch_no, item_name ", db, adOpenStatic, adLockReadOnly, adCmdText
    'rststock.Open "SELECT * FROM RTRXFILE WHERE ITEM_NAME Like '%" & TXTITEM.text & "%' AND bal_qty >0 and ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY trx_type, trx_year, vch_no, item_name ", db, adOpenStatic, adLockReadOnly, adCmdText
    rststock.Open "SELECT * FROM RTRXFILE LEFT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE RTRXFILE.ITEM_NAME Like '%" & TXTITEM.text & "%' AND RTRXFILE.BAL_QTY >0 and ucase(RTRXFILE.CATEGORY) <> 'SERVICE CHARGE' AND (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') ", db, adOpenStatic, adLockReadOnly
    Do Until rststock.EOF
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = IIf(IsNull(rststock!ITEM_CODE), "", rststock!ITEM_CODE)
        GRDSTOCK.TextMatrix(i, 2) = IIf(IsNull(rststock!ITEM_NAME), "", rststock!ITEM_NAME)
        GRDSTOCK.TextMatrix(i, 3) = ""
        GRDSTOCK.TextMatrix(i, 4) = IIf(IsNull(rststock!BAL_QTY), "", rststock!BAL_QTY)
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!ITEM_COST), "", rststock!ITEM_COST)
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!SALES_TAX), "", rststock!SALES_TAX)
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", rststock!P_RETAIL)
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTITEMMAST
            If Not (.EOF And .BOF) Then
                If Val(GRDSTOCK.TextMatrix(i, 6)) = 0 Then
                    GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(RSTITEMMAST!SALES_TAX), "", RSTITEMMAST!SALES_TAX)
                End If
                If Val(GRDSTOCK.TextMatrix(i, 7)) = 0 Then
                    GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(RSTITEMMAST!P_RETAIL), "", RSTITEMMAST!P_RETAIL)
                End If
            End If
        End With
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        GRDSTOCK.TextMatrix(i, 8) = IIf(IsDate(rststock!VCH_DATE), Format(rststock!VCH_DATE, "DD/MM/YYYY"), "")
        GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!PINV), "", rststock!PINV)
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!VCH_DESC), "", Mid(rststock!VCH_DESC, 15))
        GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!LOOSE_PACK) Or rststock!LOOSE_PACK = 0, "1", rststock!LOOSE_PACK)
        GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!VCH_NO), "", rststock!VCH_NO)
        GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!LINE_NO), "", rststock!LINE_NO)
        GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!TRX_TYPE), "", rststock!TRX_TYPE)
        GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!TRX_YEAR), "", rststock!TRX_YEAR)
        GRDSTOCK.TextMatrix(i, 16) = IIf(IsNull(rststock!PACK_TYPE), "Nos", rststock!PACK_TYPE)
        i = i + 1
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    TxtBillName.text = "Cash"
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdMake_Click()
    Dim i As Long
    Dim n As Integer
    'grdsales.FixedRows = 0
    'grdsales.Rows = 1
    i = 1
    'Screen.MousePointer = vbHourglass
    'For N = 1 To GRDSTOCK.Rows - 1
        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) > 0 Then
            If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) = 0 Then
                MsgBox "The item " & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2)) & " cannot be added since the rate is zero", vbOKOnly, "EzBiz"
            Else
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                grdsales.TextMatrix(grdsales.rows - 1, 0) = grdsales.rows - 1
                grdsales.TextMatrix(grdsales.rows - 1, 1) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) ' ITEM CODE
                grdsales.TextMatrix(grdsales.rows - 1, 2) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2) ' ITEM DESCRIPTION
                grdsales.TextMatrix(grdsales.rows - 1, 3) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3) 'QTY
                grdsales.TextMatrix(grdsales.rows - 1, 5) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6) 'TAX
                grdsales.TextMatrix(grdsales.rows - 1, 6) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7) 'NET RATE
                If MDIMAIN.lblkfc.Caption = "Y" And IsDate(MDIMAIN.DTKFCSTART.Value) And IsDate(MDIMAIN.DTKFCEND.Value) Then
                    If DateValue(TXTINVDATE.text) >= DateValue(MDIMAIN.DTKFCSTART.Value) And DateValue(TXTINVDATE.text) <= DateValue(MDIMAIN.DTKFCEND.Value) Then
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) = 12 Or Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) = 18 Or Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) = 28 Then
                            grdsales.TextMatrix(grdsales.rows - 1, 15) = 1
                        Else
                            grdsales.TextMatrix(grdsales.rows - 1, 15) = 0
                        End If
                    End If
                Else
                    grdsales.TextMatrix(grdsales.rows - 1, 15) = 0
                End If
                If MDIMAIN.LblKFCNet.Caption = "Y" Then
                    'TXTRETAILNOTAX.Text = (Val(txtNetrate.Text) - Val(TxtCessAmt.Text)) / (1 + ((Val(TXTTAX.Text) + Val(TxtKFC.Caption)) / 100) + (Val(TxtCessPer.Text) / 100))
                    grdsales.TextMatrix(grdsales.rows - 1, 4) = Round((Val(grdsales.TextMatrix(grdsales.rows - 1, 6))) / (1 + ((Val(grdsales.TextMatrix(grdsales.rows - 1, 5)) + Val(grdsales.TextMatrix(grdsales.rows - 1, 15))) / 100)), 4)
                    grdsales.TextMatrix(grdsales.rows - 1, 6) = Round(Val(grdsales.TextMatrix(grdsales.rows - 1, 4)) + (Val(grdsales.TextMatrix(grdsales.rows - 1, 4)) * Val(grdsales.TextMatrix(grdsales.rows - 1, 5)) / 100), 4)
                Else
                    grdsales.TextMatrix(grdsales.rows - 1, 4) = Round(Val(grdsales.TextMatrix(grdsales.rows - 1, 6)) * 100 / ((Val(grdsales.TextMatrix(grdsales.rows - 1, 5))) + 100), 4)
                End If
                grdsales.TextMatrix(grdsales.rows - 1, 7) = IIf(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) = 0, 1, Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11))) 'PACK
                
                'grdsales.TextMatrix(grdsales.Rows - 1, 8) = Val(grdsales.TextMatrix(grdsales.Rows - 1, 6)) * Val(grdsales.TextMatrix(grdsales.Rows - 1, 3)) 'TOTAL
                
                grdsales.TextMatrix(grdsales.rows - 1, 9) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) '"VCH_NO"
                grdsales.TextMatrix(grdsales.rows - 1, 10) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) '"LINE_NO"
                grdsales.TextMatrix(grdsales.rows - 1, 11) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14) ' TYPE
                grdsales.TextMatrix(grdsales.rows - 1, 12) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) 'YEAR
                grdsales.TextMatrix(grdsales.rows - 1, 13) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5) 'cost
                grdsales.TextMatrix(grdsales.rows - 1, 14) = GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) 'uom
                grdsales.TextMatrix(grdsales.rows - 1, 8) = Round((Val(grdsales.TextMatrix(grdsales.rows - 1, 4)) + (Val(grdsales.TextMatrix(grdsales.rows - 1, 4)) * (Val(grdsales.TextMatrix(grdsales.rows - 1, 5)) + Val(grdsales.TextMatrix(grdsales.rows - 1, 15))) / 100)) * Val(grdsales.TextMatrix(grdsales.rows - 1, 3)), 3)
                i = i + 1
            End If
        End If
    'Next N
    
    
    
    Call Calculate_Total
    Screen.MousePointer = vbNormal
    Exit Sub
        
End Sub

Private Sub CmdPrint_Click()
    If grdsales.rows <= 1 Then Exit Sub
    db.Execute "delete from TEMPTRXFILE"
    Dim RSTUNBILL As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i As Long
    Dim BILL_NUM As Double
    Dim Small_Print As Boolean
    Small_Print = True
    
    
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BILL_NUM = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From TEMPTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        Set RSTUNBILL = New ADODB.Recordset
        RSTUNBILL.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 13) & "' AND UN_BILL = 'Y'", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTUNBILL
            If Not (.EOF And .BOF) Then
                RSTUNBILL.Close
                Set RSTUNBILL = Nothing
                GoTo SKIP_UNBILL
            End If
        End With
        RSTUNBILL.Close
        Set RSTUNBILL = Nothing
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "HI"
        'RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!Category = ""
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 5)
        RSTTRXFILE!UNIT = 1
        RSTTRXFILE!VCH_DESC = "" '"Issued to     " & Trim(DataList2.Text)
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = "V"
        RSTTRXFILE!MFGR = ""
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 8)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0

        RSTTRXFILE!RETAILER_PRICE = 0
        RSTTRXFILE!CESS_PER = 0
        RSTTRXFILE!CESS_AMT = 0
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 5)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!SALE_1_FLAG = ""
        RSTTRXFILE!COM_AMT = 0
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!WARRANTY_TYPE = ""
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 14)
        RSTTRXFILE!LOOSE_FLAG = "F"
        
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "130000"
        RSTTRXFILE.Update
SKIP_UNBILL:
    Next i
    
    
    'Call ReportGeneratION_vpestimate
    Screen.MousePointer = vbHourglass
    Sleep (150)
    
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompAddress5, CompTin, CompCST, BIL_PRE, BILL_SUF, DL, ML, DL1, DL2, INV_TERMS, INV_MSG, BANK_DET, PAN_NO As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        CompAddress5 = IIf(IsNull(RSTCOMPANY!TEL_NO) Or RSTCOMPANY!TEL_NO = "", "", "Ph: " & RSTCOMPANY!TEL_NO)
        CompAddress3 = IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", "Ph: " & RSTCOMPANY!FAX_NO)
        CompAddress4 = IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!EMAIL_ADD)
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
        DL = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "DL No. " & RSTCOMPANY!DL_NO)
        ML = IIf(IsNull(RSTCOMPANY!ML_NO) Or RSTCOMPANY!DL_NO = "", "", "ML No. " & RSTCOMPANY!ML_NO)
        BIL_PRE = IIf(IsNull(RSTCOMPANY!PREFIX_8V), "", RSTCOMPANY!PREFIX_8V)
        BILL_SUF = IIf(IsNull(RSTCOMPANY!SUFIX_8V), "", RSTCOMPANY!SUFIX_8V)
        INV_TERMS = IIf(IsNull(RSTCOMPANY!INV_TERMS) Or RSTCOMPANY!INV_TERMS = "", "", RSTCOMPANY!INV_TERMS)
        INV_MSG = IIf(IsNull(RSTCOMPANY!INV_MSGS) Or RSTCOMPANY!INV_MSGS = "", "", RSTCOMPANY!INV_MSGS)
        BANK_DET = IIf(IsNull(RSTCOMPANY!bank_details) Or RSTCOMPANY!bank_details = "", "", RSTCOMPANY!bank_details)
        PAN_NO = IIf(IsNull(RSTCOMPANY!PAN_NO) Or RSTCOMPANY!PAN_NO = "", "", RSTCOMPANY!PAN_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
        
                
    lblnetamount.Tag = Round(Val(lblnetamount.Caption), 2)
    If Val(MDIMAIN.StatusBar.Panels(11).text) = 1 Then
        If Small_Print = True Then
            'ReportNameVar = Rptpath & "rptbillretail"
            ReportNameVar = Rptpath & "RPTGSTBILLA51"
        Else
            ReportNameVar = Rptpath & "RPTGSTBILL1"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Set CRXFormulaFields = Report.FormulaFields
        
        For i = 1 To Report.Database.Tables.COUNT
            Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        Next i
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM TEMPTRXFILE ")
            Report.Database.SetDataSource oRs, 3, 1
            Set oRs = Nothing
            
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM ITEMMAST ")
            Report.Database.SetDataSource oRs, 3, 2
            Set oRs = Nothing
        End If
        Report.DiscardSavedData
        Report.VerifyOnEveryPrint = True
        Set CRXFormulaFields = Report.FormulaFields
        For Each CRXFormulaField In CRXFormulaFields
            If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
            If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
            If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
            If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
            If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
            If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
            If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.text = "'" & CompAddress5 & "'"
            If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
            If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
            If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.text = "'" & DL & "'"
            If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.text = "'" & ML & "'"
            If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
            If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.text = "'" & INV_MSG & "'"
            If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
            If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
            If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.text = "'" & DL1 & "'"
            If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.text = "'" & DL2 & "'"
            If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TxtBillName.text & "'"
            If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TxtBillName.text) & "'"
            If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
            If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
            
            'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
            'If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
    '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
            If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
    '        If Tax_Print = False Then
    '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
    '        End If
            If CRXFormulaField.Name = "{@VCH_NO}" Then
                Me.Tag = BIL_PRE & Format(BILL_NUM, bill_for) & BILL_SUF
                CRXFormulaField.text = "'" & Me.Tag & "' "
            End If
            'If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
    '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
            'If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
            If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
        Next
    Else
        If Small_Print = True Then
            'ReportNameVar = Rptpath & "rptbillretail"
            ReportNameVar = Rptpath & "RPTGSTBILLA5"
        Else
            ReportNameVar = Rptpath & "rptGSTBILL"
        End If
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Set CRXFormulaFields = Report.FormulaFields
        
        For i = 1 To Report.OpenSubreport("RPTBILL1.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL1.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To Report.OpenSubreport("RPTBILL2.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL2.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To Report.OpenSubreport("RPTBILL3.rpt").Database.Tables.COUNT
            Report.OpenSubreport("RPTBILL3.rpt").Database.Tables(i).SetLogOnInfo strConnection
        Next i
        For i = 1 To 3
            'Set CRXFormulaFields = Report.FormulaFields
            Set CRXFormulaFields = Report.OpenSubreport("RPTBILL" & i & ".rpt").FormulaFields
            For Each CRXFormulaField In CRXFormulaFields
                If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
                If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
                If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
                If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
                If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
                If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
                If CRXFormulaField.Name = "{@Comp_Address5}" Then CRXFormulaField.text = "'" & CompAddress5 & "'"
                If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
                If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
                If CRXFormulaField.Name = "{@DL}" Then CRXFormulaField.text = "'" & DL & "'"
                If CRXFormulaField.Name = "{@ML}" Then CRXFormulaField.text = "'" & ML & "'"
                If CRXFormulaField.Name = "{@inv_terms}" Then CRXFormulaField.text = "'" & INV_TERMS & "'"
                If CRXFormulaField.Name = "{@inv_msg}" Then CRXFormulaField.text = "'" & INV_MSG & "'"
                If CRXFormulaField.Name = "{@bank}" Then CRXFormulaField.text = "'" & BANK_DET & "'"
                If CRXFormulaField.Name = "{@pan}" Then CRXFormulaField.text = "'" & PAN_NO & "'"
                If CRXFormulaField.Name = "{@DL1}" Then CRXFormulaField.text = "'" & DL1 & "'"
                If CRXFormulaField.Name = "{@DL2}" Then CRXFormulaField.text = "'" & DL2 & "'"
                If CRXFormulaField.Name = "{@Company}" Then CRXFormulaField.text = "'" & TxtBillName.text & "'"
                If CRXFormulaField.Name = "{@CustName}" Then CRXFormulaField.text = "'" & Trim(TxtBillName.text) & "'"
                If CRXFormulaField.Name = "{@CustAddress}" Then CRXFormulaField.text = "'" & Trim(TxtBillAddress.text) & "'"
                If CRXFormulaField.Name = "{@SR}" Then CRXFormulaField.text = " 'N' "
                'If CRXFormulaField.Name = "{@TOF}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLFOT.Caption), 2), "0.00") & "'"
                'If CRXFormulaField.Name = "{@Disc}" Then CRXFormulaField.Text = "'" & Format(Round(Val(LBLDISCAMT.Caption), 2), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round1}" Then CRXFormulaField.Text = "'" & Format(Val(lblnetamount.Tag), "0.00") & "'"
        '            If CRXFormulaField.Name = "{@Round2}" Then CRXFormulaField.Text = "'" & Format(Val(Round(Val(LBLTOTAL.Caption) + Val(LBLFOT.Caption) - Val(LBLDISCAMT.Caption), 0)), "0.00") & "'"
                If CRXFormulaField.Name = "{@Total}" Then CRXFormulaField.text = "'" & Format(Val(LBLTOTAL.Caption), "0.00") & "'"
        '        If Tax_Print = False Then
        '            If CRXFormulaField.Name = "{@Figure}" Then CRXFormulaField.Text = "'" & Trim(LBLFOT.Tag) & "'"
        '        End If
                If CRXFormulaField.Name = "{@VCH_NO}" Then
                    Me.Tag = BIL_PRE & Format(BILL_NUM, bill_for) & BILL_SUF
                    CRXFormulaField.text = "'" & Me.Tag & "' "
                End If
                'If CRXFormulaField.Name = "{@DISCAMT}" Then CRXFormulaField.Text = " " & Val(LBLDISCAMT.Caption) & " "
        '            If CRXFormulaField.Name = "{@NetGrandTotal}" Then CRXFormulaField.Text = "'" & Format(Round(Val(lblnetamount.Caption), 0), "0.00") & "'"
                'If CRXFormulaField.Name = "{@CUSTCODE}" Then CRXFormulaField.Text = "'" & Trim(TxtCode.Text) & "'"
                If CRXFormulaField.Name = "{@Credit}" Then CRXFormulaField.text = "'Cash'"
            Next
        Next i
    End If
    If MDIMAIN.StatusBar.Panels(13).text = "Y" Then
        'Preview
        frmreport.Caption = "BILL"
        Call GENERATEREPORT
        Screen.MousePointer = vbNormal
    Else
        '    '''No Preview
        Report.PrintOut (False)
        Set CRXFormulaFields = Nothing
        Set CRXFormulaField = Nothing
        Set crxApplication = Nothing
        Set Report = Nothing
    End If
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub CmdSave_Click()
    If grdsales.rows <= 1 Then Exit Sub
    If Val(lblnetamount.Caption) = 0 Then
        MsgBox "Amount is Zero", vbOKOnly, "Easy Bill"
        Exit Sub
    End If
    Call Make_Invoice
    Call Calculate_sales
    Call Sales_Register_Sum
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Shift = vbCtrlMask Then
        Select Case KeyCode
            Case vbKeyG
                'CmdMake_Click
                'If MsgBox("Are You Sure you want to generate the Invoice?", vbYesNo, "EzBiz") = vbNo Then Exit Sub
                CmdSave_Click
                If SAVE_BILL = True Then
                    If MsgBox("Are You Sure you want to Print the Invoice?", vbYesNo + vbDefaultButton2, "EzBiz") = vbYes Then CmdPrint_Click
                    Call CmDDisplay_Click
                    grdsales.FixedRows = 0
                    grdsales.rows = 1
                    TxtBillName.text = "Cash"
                    LBLTOTAL.Caption = ""
                    lblnetamount.Caption = ""
                    LBLTOTALCOST.Caption = ""
                    lblselAMT.Caption = ""
                    LBLBILLNO.Caption = Val(LBLBILLNO.Caption) + 1
                    GRDSTOCK.Col = 3
                    GRDSTOCK.SetFocus
                End If
        End Select
    Else
        Select Case KeyCode
            Case vbKeyEscape
                TXTITEM.SetFocus
        End Select
    End If
    
    
    
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.SetFocus
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                TxtBillName.SetFocus
            End If
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Form_Load()
    GRDSTOCK.TextMatrix(0, 0) = "Sl"
    GRDSTOCK.TextMatrix(0, 1) = "Item Code"
    GRDSTOCK.TextMatrix(0, 2) = "Item Description"
    GRDSTOCK.TextMatrix(0, 3) = "Bill Qty"
    GRDSTOCK.TextMatrix(0, 4) = "Bal Qty"
    GRDSTOCK.TextMatrix(0, 5) = "Cost"
    GRDSTOCK.TextMatrix(0, 6) = "Tax"
    GRDSTOCK.TextMatrix(0, 7) = "Ret. Price"
    GRDSTOCK.TextMatrix(0, 8) = "Inv Date"
    GRDSTOCK.TextMatrix(0, 9) = "Inv No."
    GRDSTOCK.TextMatrix(0, 10) = "Supplier"
    GRDSTOCK.TextMatrix(0, 11) = "" '"Pack"
    GRDSTOCK.TextMatrix(0, 12) = "" '"VCH_NO"
    GRDSTOCK.TextMatrix(0, 13) = "" '"LINE_NO"
    GRDSTOCK.TextMatrix(0, 14) = "" ' TYPE
    GRDSTOCK.TextMatrix(0, 15) = "" 'YEAR
    GRDSTOCK.TextMatrix(0, 16) = "" '"UOM" 'packtype

    GRDSTOCK.ColWidth(0) = 600
    GRDSTOCK.ColWidth(1) = 0
    GRDSTOCK.ColWidth(2) = 2500
    GRDSTOCK.ColWidth(3) = 900
    GRDSTOCK.ColWidth(4) = 800
    GRDSTOCK.ColWidth(5) = 1000
    GRDSTOCK.ColWidth(6) = 1000
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(8) = 1100
    GRDSTOCK.ColWidth(9) = 1000
    GRDSTOCK.ColWidth(10) = 1800
    GRDSTOCK.ColWidth(11) = 0
    GRDSTOCK.ColWidth(12) = 0
    GRDSTOCK.ColWidth(13) = 0
    GRDSTOCK.ColWidth(14) = 0
    GRDSTOCK.ColWidth(15) = 0
    GRDSTOCK.ColWidth(16) = 0
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 4
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 1
    GRDSTOCK.ColAlignment(11) = 4
    GRDSTOCK.ColAlignment(12) = 4
    GRDSTOCK.ColAlignment(13) = 4
    GRDSTOCK.ColAlignment(14) = 4
    GRDSTOCK.ColAlignment(15) = 4
    GRDSTOCK.ColAlignment(16) = 4
    
    
    grdsales.TextMatrix(0, 0) = "Sl"
    grdsales.TextMatrix(0, 1) = "Item Code"
    grdsales.TextMatrix(0, 2) = "Item Description"
    grdsales.TextMatrix(0, 3) = "Qty"
    grdsales.TextMatrix(0, 4) = "Rate"
    grdsales.TextMatrix(0, 5) = "Tax"
    grdsales.TextMatrix(0, 6) = "Net Rate"
    grdsales.TextMatrix(0, 7) = "Pack"
    grdsales.TextMatrix(0, 8) = "Total"
    grdsales.TextMatrix(0, 9) = "" '"VCH_NO"
    grdsales.TextMatrix(0, 10) = "" '"LINE_NO"
    grdsales.TextMatrix(0, 11) = "" '"TRX_TYPE"
    grdsales.TextMatrix(0, 12) = "" '"TRX_YEAR"
    grdsales.TextMatrix(0, 13) = "" '"cost"
    grdsales.TextMatrix(0, 14) = "" '"packtype"
    
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2000
    grdsales.ColWidth(3) = 700
    grdsales.ColWidth(4) = 900
    grdsales.ColWidth(5) = 600
    grdsales.ColWidth(6) = 900
    grdsales.ColWidth(7) = 800
    grdsales.ColWidth(8) = 1000
    grdsales.ColWidth(9) = 0
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 0
    grdsales.ColWidth(13) = 0
    grdsales.ColWidth(14) = 0
    
    grdsales.ColAlignment(0) = 1
    grdsales.ColAlignment(1) = 1
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 4
    grdsales.ColAlignment(4) = 7
    grdsales.ColAlignment(5) = 4
    grdsales.ColAlignment(6) = 7
    grdsales.ColAlignment(7) = 4
    grdsales.ColAlignment(8) = 7
    grdsales.ColAlignment(9) = 4
    grdsales.ColAlignment(10) = 4
    grdsales.ColAlignment(11) = 4
    grdsales.ColAlignment(12) = 4
        
    'TXTINVDATE.Text = Format(Date, "dd/mm/yyyy")
    TxtBillName.text = "Cash"
    On Error GoTo ErrHand
    Dim rstBILL  As ADODB.Recordset
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI'", db, adOpenStatic, adLockReadOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        LBLBILLNO.Caption = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
    
    Left = 0
    Top = 0
    'Height = 10000
    'Width = 12840
    Call CmDDisplay_Click
    Call Calculate_sales
    Call Sales_Register_Sum
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If REPFLAG = False Then RSTREP.Close
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 6, 7, 3
                    On Error Resume Next
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop '+ 350
                    TXTsample.Left = GRDSTOCK.CellLeft '+ 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
            End Select
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
End Sub

Private Sub TXTINVDATE_LostFocus()
    Call Calculate_sales
    Call Sales_Register_Sum
End Sub

Private Sub TXTITEM_Change()
    Call CmDDisplay_Click
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyDown
            If GRDSTOCK.rows <= 1 Then Exit Sub
            GRDSTOCK.Col = 3
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(TXTINVDATE.text) Then
                MsgBox "Select proper invoice date", vbOKOnly, "EzBiz"
                TXTINVDATE.SetFocus
                Exit Sub
            End If
            Select Case GRDSTOCK.Col
                Case 7 'Rate
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    Call Total_Amount
                    GRDSTOCK.SetFocus
                Case 6 'TAX
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    Call Total_Amount
                    GRDSTOCK.SetFocus
                Case 3   'Qty
                    If Val(TXTsample.text) > Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4)) Then
                        MsgBox "Item is greater than the available qty", vbOKOnly, "EzBiz"
                        TXTsample.SetFocus
                        Exit Sub
                    End If
                    If Val(TXTsample.text) > 0 Then
                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 7)) = 0 Then
                            MsgBox "Please enter Selling Price", vbOKOnly, "EzBiz"
                        End If
'                        If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 6)) = 0 Then
'                            If MsgBox("The Tax Rate is Zero. Are You Sure?", vbYesNo, "EzBiz") = vbNo Then Exit Sub
'                        End If
                    End If
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    Call Total_Amount
                    Call CmdMake_Click
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
        Case vbKeyUp, vbKeyDown
            Call TXTsample_KeyDown(13, 0)
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 6, 7, 3
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Function Total_Amount()
    lblselAMT.Caption = ""
    Dim i As Long
    For i = 1 To GRDSTOCK.rows - 1
        lblselAMT.Caption = Val(lblselAMT.Caption) + (Val(GRDSTOCK.TextMatrix(i, 7)) * Val(GRDSTOCK.TextMatrix(i, 3)))
    Next i
End Function

Private Function Make_Invoice()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTP_RATE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim rstBILL As ADODB.Recordset
    Dim i, BILL_NUM As Double
    Dim TRXVALUE As Double
    Dim DAY_DATE As String
    Dim MONTH_DATE As String
    Dim YEAR_DATE As String
    Dim E_DATE As Date
    i = 0
    On Error GoTo ErrHand
    SAVE_BILL = False
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Select proper invoice date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Function
    End If

    If Trim(TxtBillName.text) = "" Then TxtBillName.text = "Cash"
    Set rstBILL = New ADODB.Recordset
    rstBILL.Open "Select MAX(VCH_NO) From TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'HI'", db, adOpenForwardOnly
    If Not (rstBILL.EOF And rstBILL.BOF) Then
        BILL_NUM = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
    End If
    rstBILL.Close
    Set rstBILL = Nothing
        
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 1) = "" Then GoTo SKIP_2
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then GoTo SKIP_2
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                '!ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(I, 3))
                If (IsNull(!FREE_QTY)) Then !FREE_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 7))), 4)
                !FREE_QTY = 0
                !CLOSE_QTY = !CLOSE_QTY - Round(((Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 7))), 3)
    
                If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
                !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(i, 8))
                If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
                !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(i, 8))
                RSTTRXFILE.Update
            End If
        End With
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "' AND RTRXFILE.TRX_TYPE = '" & Trim(grdsales.TextMatrix(i, 11)) & "' AND RTRXFILE.VCH_NO = " & Val(grdsales.TextMatrix(i, 9)) & " AND RTRXFILE.LINE_NO = " & Val(grdsales.TextMatrix(i, 10)) & " AND RTRXFILE.TRX_YEAR = '" & Val(grdsales.TextMatrix(i, 12)) & "' AND BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
                If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
                !ISSUE_QTY = !ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 7)), 3)
                !BAL_QTY = !BAL_QTY - Round((Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 7)), 3)
                RSTTRXFILE.Update
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            Else
                'BALQTY = 0
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "' AND BAL_QTY > 0 ORDER BY BAL_QTY DESC", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    If (IsNull(RSTTRXFILE!ISSUE_QTY)) Then RSTTRXFILE!ISSUE_QTY = 0
                    If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                    'BALQTY = RSTTRXFILE!BAL_QTY
                    RSTTRXFILE!ISSUE_QTY = RSTTRXFILE!ISSUE_QTY + Round((Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 7)), 3)
                    RSTTRXFILE!BAL_QTY = RSTTRXFILE!BAL_QTY - Round((Val(grdsales.TextMatrix(i, 3))) * Val(grdsales.TextMatrix(i, 7)), 3)
                    
                    grdsales.TextMatrix(i, 9) = RSTTRXFILE!VCH_NO
                    grdsales.TextMatrix(i, 10) = RSTTRXFILE!LINE_NO
                    grdsales.TextMatrix(i, 11) = RSTTRXFILE!TRX_TYPE
                    grdsales.TextMatrix(i, 11) = RSTTRXFILE!TRX_YEAR
        
                    RSTTRXFILE.Update
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
        End With
SKIP_2:
    Next i
    
    db.Execute "delete FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & ""
    db.Execute "delete FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & ""
    db.Execute "delete FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & "", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!TRX_TYPE = "HI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
        RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = "130000"
        RSTTRXFILE!ACT_NAME = "CASH"
        RSTTRXFILE!DISCOUNT = 0
    End If
    
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!NET_AMOUNT = Val(lblnetamount.Caption)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = "130000"
    RSTTRXFILE!ACT_NAME = "CASH"
    RSTTRXFILE!DISCOUNT = 0
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!ROUNDED_OFF = 0
    RSTTRXFILE!PAY_AMOUNT = Val(LBLTOTALCOST.Caption)
    RSTTRXFILE!ADD_AMOUNT = 0
    RSTTRXFILE!REF_NO = ""
    RSTTRXFILE!SLSM_CODE = "P"
    RSTTRXFILE!DISCOUNT = 0
    RSTTRXFILE!CHECK_FLAG = "I"
    RSTTRXFILE!POST_FLAG = "Y"
    RSTTRXFILE!CFORM_NO = Time
    RSTTRXFILE!REMARKS = Trim(TxtBillName.text)
    RSTTRXFILE!DISC_PERS = 0
    RSTTRXFILE!AST_PERS = 0
    RSTTRXFILE!AST_AMNT = 0
    RSTTRXFILE!BANK_CHARGE = 0
    RSTTRXFILE!VEHICLE = ""
    RSTTRXFILE!PHONE = ""
    RSTTRXFILE!TIN = ""
    RSTTRXFILE!FRIEGHT = 0
    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Date
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!cr_days = 0
    RSTTRXFILE!BILL_NAME = Trim(TxtBillName.text)
    RSTTRXFILE!BILL_ADDRESS = Trim(TxtBillAddress.text)
    RSTTRXFILE!AGENT_CODE = ""
    RSTTRXFILE!AGENT_NAME = ""
    RSTTRXFILE!BILL_TYPE = "R"
    RSTTRXFILE!CN_REF = Null
    
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXSUB WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & "", db, adOpenStatic, adLockOptimistic, adCmdText
    
    'grdsales.TextMatrix(I, 15) = Trim(TXTTRXTYPE.Text)
    
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 1) = "" Then GoTo SKIP_3
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then GoTo SKIP_3
        RSTTRXFILE.AddNew
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!TRX_TYPE = "HI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!R_VCH_NO = IIf(grdsales.TextMatrix(i, 9) = "", 0, grdsales.TextMatrix(i, 9))
        RSTTRXFILE!R_LINE_NO = IIf(grdsales.TextMatrix(i, 10) = "", 0, grdsales.TextMatrix(i, 10))
        RSTTRXFILE!R_TRX_TYPE = IIf(grdsales.TextMatrix(i, 11) = "", "MI", grdsales.TextMatrix(i, 11))
        RSTTRXFILE!R_TRX_YEAR = IIf(grdsales.TextMatrix(i, 12) = "", Year(MDIMAIN.DTFROM.Value), grdsales.TextMatrix(i, 12))
        RSTTRXFILE!QTY = grdsales.TextMatrix(i, 3)
        RSTTRXFILE.Update
SKIP_3:
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * FROM TRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE='HI' AND VCH_NO = " & BILL_NUM & "", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 1) = "" Then GoTo SKIP_4
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then GoTo SKIP_4
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "HI"
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_NO = BILL_NUM
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = i
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        With RSTITEMMAST
            If Not (.EOF And .BOF) Then
                RSTTRXFILE!Category = RSTITEMMAST!Category
                RSTTRXFILE!MFGR = RSTITEMMAST!MANUFACTURER
            End If
        End With
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        RSTTRXFILE!ITEM_CODE = grdsales.TextMatrix(i, 1)
        RSTTRXFILE!ITEM_NAME = grdsales.TextMatrix(i, 2)
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(i, 13))
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!P_RETAIL = Val(grdsales.TextMatrix(i, 6))
        RSTTRXFILE!P_RETAILWOTAX = Val(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!COM_AMT = 0
        RSTTRXFILE!COM_FLAG = "N"
        RSTTRXFILE!LOOSE_FLAG = "F"
        RSTTRXFILE!LOOSE_PACK = Val(grdsales.TextMatrix(i, 7))
        RSTTRXFILE!SALES_TAX = grdsales.TextMatrix(i, 5)
        RSTTRXFILE!UNIT = 1
        RSTTRXFILE!VCH_DESC = "Issued to     " & Mid(Trim(TxtBillName.text), 1, 30)
        RSTTRXFILE!REF_NO = ""
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CHECK_FLAG = "V"
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = 0
        RSTTRXFILE!kfc_tax = Val(grdsales.TextMatrix(i, 15))
        RSTTRXFILE!TRX_TOTAL = grdsales.TextMatrix(i, 8)
        RSTTRXFILE!LINE_DISC = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!MODIFY_DATE = Date
        RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "130000"
        RSTTRXFILE!SALE_1_FLAG = ""
        RSTTRXFILE!WARRANTY = Null
        RSTTRXFILE!WARRANTY_TYPE = ""
        RSTTRXFILE!PACK_TYPE = grdsales.TextMatrix(i, 14)
        RSTTRXFILE.Update
SKIP_4:
    Next i
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim Max_No As Long
    Max_No = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        Max_No = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & BILL_NUM & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'HI'"
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & BILL_NUM & " AND INV_TYPE = 'RT' AND INV_TRX_TYPE = 'HI'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!REC_NO = Max_No + 1
        RSTITEMMAST!INV_TYPE = "RT"
        RSTITEMMAST!INV_TRX_TYPE = "HI"
        RSTITEMMAST!INV_NO = BILL_NUM
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    End If
    'If lblcredit.Caption <> "0" Then
    RSTITEMMAST!AMOUNT = Val(lblnetamount.Caption)
    RSTITEMMAST!TRX_TYPE = "CR"
    RSTITEMMAST!CHECK_FLAG = "S"
    RSTITEMMAST!ACT_CODE = 130000
    RSTITEMMAST!ACT_NAME = "CASH"
    RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
    db.CommitTrans
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    SAVE_BILL = True
    MsgBox "Success. Bill No. " & BILL_NUM & " Generated", vbOKOnly, "EzBiz"
SKIP:
    Exit Function
ErrHand:
    MsgBox err.Description
End Function

Private Function Calculate_Total()
    LBLTOTAL.Caption = ""
    lblnetamount.Caption = ""
    LBLTOTALCOST.Caption = ""
    Dim i As Long
    For i = 1 To grdsales.rows - 1
        If grdsales.TextMatrix(i, 1) = "" Then GoTo SKIP_2
        If Val(grdsales.TextMatrix(i, 3)) = 0 Then GoTo SKIP_2
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + (Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 4)))
        lblnetamount.Caption = Val(lblnetamount.Caption) + Val(grdsales.TextMatrix(i, 8)) '(Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 6)))
        LBLTOTALCOST.Caption = Val(LBLTOTALCOST.Caption) + (Val(grdsales.TextMatrix(i, 3)) * Val(grdsales.TextMatrix(i, 13)))
SKIP_2:
    Next i
    LBLTOTAL.Caption = Format(Round(Val(LBLTOTAL.Caption), 2), "0.00")
    lblnetamount.Caption = Format(Round(Val(lblnetamount.Caption), 2), "0.00")
    LBLTOTALCOST.Caption = Format(Round(Val(LBLTOTALCOST.Caption), 2), "0.00")
    
End Function

Private Function Calculate_sales()
    Dim rstTRANX As ADODB.Recordset
    Dim TOT_SALE As Long
    Dim FROM_DATE As Date
    Dim TO_DATE As Date
    Dim D As String
    
    LBLsalesday.Caption = "0.00"
    lblsalesmonth.Caption = "0.00"
    lblpurcahsemonth.Caption = "0.00"
    If IsDate(TXTINVDATE.text) Then
        FROM_DATE = "01/" & Month(TXTINVDATE.text) & "/" & Year(TXTINVDATE.text)
        D = Format(LastDayOfMonth(TXTINVDATE.text), "00")
        TO_DATE = D & "/" & Month(TXTINVDATE.text) & "/" & Year(TXTINVDATE.text)
                   
        On Error GoTo ErrHand
        Screen.MousePointer = vbHourglass
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT * From TRXMAST WHERE TRX_TYPE='HI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'        Do Until rstTRANX.EOF
'            TOT_SALE = TOT_SALE + IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
'            rstTRANX.MoveNext
'        Loop
'        rstTRANX.Close
'        Set rstTRANX = Nothing
        
        TOT_SALE = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            TOT_SALE = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        lblsalesmonth.Caption = Format(TOT_SALE, "0.00")
        
        TOT_SALE = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(NET_AMOUNT) FROM TRANSMAST WHERE TRX_TYPE='PI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            TOT_SALE = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        lblpurcahsemonth.Caption = Format(TOT_SALE, "0.00")
        
'        TOT_SALE = 0
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT * From TRANSMAST WHERE TRX_TYPE='PI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'        Do Until rstTRANX.EOF
'            TOT_SALE = TOT_SALE + IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
'            rstTRANX.MoveNext
'        Loop
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'        lblpurcahsemonth.Caption = Format(TOT_SALE, "0.00")
    
        TOT_SALE = 0
        Set rstTRANX = New ADODB.Recordset
        rstTRANX.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE VCH_DATE = '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
        If Not (rstTRANX.EOF And rstTRANX.BOF) Then
            TOT_SALE = Format(IIf(IsNull(rstTRANX.Fields(0)), 0, rstTRANX.Fields(0)), "0.00")
        End If
        rstTRANX.Close
        Set rstTRANX = Nothing
        LBLsalesday.Caption = Format(TOT_SALE, "0.00")
        
'        TOT_SALE = 0
'        Set rstTRANX = New ADODB.Recordset
'        rstTRANX.Open "SELECT * From TRXMAST WHERE TRX_TYPE='HI' AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_DATE = '" & Format(TXTINVDATE.text, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'        Do Until rstTRANX.EOF
'            TOT_SALE = TOT_SALE + IIf(IsNull(rstTRANX!NET_AMOUNT), 0, rstTRANX!NET_AMOUNT)
'            rstTRANX.MoveNext
'        Loop
'        rstTRANX.Close
'        Set rstTRANX = Nothing
'        LBLsalesday.Caption = Format(TOT_SALE, "0.00")
    End If
    'LBLRETURNED.Caption = Format(TOT_RET, "0.00")
    Screen.MousePointer = vbNormal
    Exit Function
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub grdsales_Click()
    txtsales.Visible = False
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    If grdsales.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 3
                    On Error Resume Next
                    txtsales.Visible = True
                    txtsales.Top = grdsales.CellTop '+ 350
                    txtsales.Left = grdsales.CellLeft '+ 50
                    txtsales.Width = grdsales.CellWidth
                    txtsales.Height = grdsales.CellHeight
                    txtsales.text = grdsales.TextMatrix(grdsales.Row, grdsales.Col)
                    txtsales.SetFocus
            End Select
    End Select
End Sub

Private Sub grdsales_Scroll()
    txtsales.Visible = False
End Sub

Private Sub TxtSales_GotFocus()
    txtsales.SelStart = 0
    txtsales.SelLength = Len(txtsales.text)
End Sub

Private Sub TxtSales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim M_STOCK As Double
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 3   'Qty
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Format(Val(txtsales.text), "0.00")
                    grdsales.TextMatrix(grdsales.Row, 8) = Val(grdsales.TextMatrix(grdsales.Row, 6)) * Val(grdsales.TextMatrix(grdsales.Row, 3)) 'TOTAL
                    grdsales.Enabled = True
                    txtsales.Visible = False
                    grdsales.SetFocus
                    Call Calculate_Total
            End Select
        Case vbKeyEscape
            txtsales.Visible = False
            grdsales.SetFocus
        Case vbKeyUp, vbKeyDown
            Call TxtSales_KeyDown(13, 0)
    End Select
        Exit Sub
ErrHand:
    MsgBox err.Description
    
End Sub

Private Sub TxtSales_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 3
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Function Sales_Register_Sum()
    Dim rstTRANX, rststock As ADODB.Recordset
    Dim i As Integer
    
    Dim FROM_DATE As Date
    Dim TO_DATE As Date
    Dim D As String
'    On Error Resume Next
'    'db.Execute "DROP TABLE TEMP_REPORT "
'    On Error GoTo eRRHAND
    
    GRDTranx.rows = 2
    GRDTranx.Cols = 1
    
    If Not IsDate(TXTINVDATE.text) Then Exit Function
    FROM_DATE = "01/" & Format(Month(TXTINVDATE.text), "00") & "/" & Format(Year(TXTINVDATE.text), "0000")
    D = Format(LastDayOfMonth(TXTINVDATE.text), "00")
    TO_DATE = Format(D, "00") & "/" & Format(Month(TXTINVDATE.text), "00") & "/" & Format(Year(TXTINVDATE.text), "0000")
        
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT SALES_TAX From TRXFILE WHERE VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    'GRDTranx.Rows = 2
    GRDTranx.Cols = rstTRANX.RecordCount  '* 2
    
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE SALES_TAX = " & rstTRANX!SALES_TAX & " and VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' AND (TRX_TYPE = 'HI' OR TRX_TYPE = 'GI') AND (ISNULL(UN_BILL) OR UN_BILL <> 'Y') ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            GRDTranx.ColWidth(i) = 1400
            GRDTranx.ColAlignment(i) = 4
            
            GRDTranx.TextMatrix(0, i) = rstTRANX!SALES_TAX & "%"
            GRDTranx.TextMatrix(1, i) = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            GRDTranx.TextMatrix(1, i) = Format(Round(GRDTranx.TextMatrix(1, i), 2), "0.00")
        End If
        rststock.Close
        Set rststock = Nothing
            
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    
    GRDTranx2.rows = 2
    GRDTranx2.Cols = 1
        
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT SALES_TAX From RTRXFILE WHERE VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI' or TRX_TYPE='LP') ORDER BY SALES_TAX", db, adOpenStatic, adLockReadOnly
    'GRDTranx2.Rows = 2
    GRDTranx2.Cols = rstTRANX.RecordCount  '* 2
    
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(TRX_TOTAL) FROM RTRXFILE WHERE SALES_TAX = " & rstTRANX!SALES_TAX & " and VCH_DATE >= '" & Format(FROM_DATE, "yyyy/mm/dd") & "' AND VCH_DATE <= '" & Format(TO_DATE, "yyyy/mm/dd") & "' AND (TRX_TYPE='PI' or TRX_TYPE='LP') ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            GRDTranx2.ColWidth(i) = 1400
            GRDTranx2.ColAlignment(i) = 4
            
            GRDTranx2.TextMatrix(0, i) = rstTRANX!SALES_TAX & "%"
            GRDTranx2.TextMatrix(1, i) = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
            GRDTranx2.TextMatrix(1, i) = Format(Round(GRDTranx2.TextMatrix(1, i), 2), "0.00")
        End If
        rststock.Close
        Set rststock = Nothing
            
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing

   
    'vbalProgressBar1.ShowText = False
    'vbalProgressBar1.value = 0
    Screen.MousePointer = vbDefault
    Exit Function
    
    
ErrHand:
    Screen.MousePointer = vbDefault
    MsgBox err.Description

End Function
