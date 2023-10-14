VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmdelitems 
   Caption         =   "Deleted Items Register"
   ClientHeight    =   7920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13155
   Icon            =   "frmdelitems.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   13155
   Begin MSDataGridLib.DataGrid grdmsc 
      Height          =   7815
      Left            =   15
      TabIndex        =   0
      Top             =   45
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   13785
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmdelitems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    On Error GoTo ErrHand
    Dim RSTTRXFILE As ADODB.Recordset
    Screen.MousePointer = vbHourglass
    Set grdmsc.DataSource = Nothing
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .Open "SELECT * FROM can_trxfile ORDER BY vch_date asc", db, adOpenForwardOnly
        Set grdmsc.DataSource = adoGrid
'
'    grdmsc.Columns(0).Caption = "ITEM CODE"
'    grdmsc.Columns(1).Caption = "ITEM NAME"
'    grdmsc.Columns(2).Caption = "QTY"
'    grdmsc.Columns(3).Caption = "UOM"
'    grdmsc.Columns(4).Caption = "MRP"
'    grdmsc.Columns(5).Caption = "R. PRICE"
'    grdmsc.Columns(6).Caption = "W. PRICE"
'    grdmsc.Columns(7).Caption = "V. PRICE"
'    grdmsc.Columns(8).Caption = "GST%"
'    grdmsc.Columns(9).Caption = "COST"
'    grdmsc.Columns(10).Caption = "NET COST"
'    grdmsc.Columns(11).Caption = "CUST DISC"
'    grdmsc.Columns(12).Caption = "LR. PRICE"
'    grdmsc.Columns(13).Caption = "LW. PRICE"
'    grdmsc.Columns(14).Caption = "CATEGORY"
'    grdmsc.Columns(15).Caption = "COMPANY"
'    grdmsc.Columns(16).Caption = "CESS%"
'    grdmsc.Columns(17).Caption = "ADDL CESS"
'
'    grdmsc.Columns(0).Width = 1200
'    grdmsc.Columns(1).Width = 6000
'    grdmsc.Columns(2).Width = 1200
'    grdmsc.Columns(3).Width = 900
'    grdmsc.Columns(4).Width = 1000
'    grdmsc.Columns(5).Width = 1000
'    grdmsc.Columns(6).Width = 1000
'    grdmsc.Columns(7).Width = 1000
'    grdmsc.Columns(8).Width = 900
'    grdmsc.Columns(9).Width = 1000
'    grdmsc.Columns(10).Width = 1000
'    grdmsc.Columns(11).Width = 1000
'    grdmsc.Columns(12).Width = 1100
'    grdmsc.Columns(13).Width = 1100
'    grdmsc.Columns(14).Width = 1500
'    grdmsc.Columns(15).Width = 1500
'    grdmsc.Columns(16).Width = 1000
'    grdmsc.Columns(17).Width = 1000
    End With
    Me.Height = 8505
    Me.Width = 13395
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub
