VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FRMBATCH 
   BackColor       =   &H000040C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6135
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataGridLib.DataGrid grdsub 
      Height          =   2910
      Left            =   75
      TabIndex        =   0
      Top             =   345
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   5133
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   4
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
      Caption         =   "MEDICINE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   240
      Index           =   0
      Left            =   3150
      TabIndex        =   2
      Top             =   105
      Width           =   2850
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
      Caption         =   "BATCH WISE LIST FOR THE ITEM "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   9
      Left            =   75
      TabIndex        =   1
      Top             =   105
      Width           =   3075
   End
End
Attribute VB_Name = "FRMBATCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean

Private Sub Form_Activate()
    FRMSALE.Enabled = False
    MDIMAIN.Enabled = False
End Sub

Private Sub Form_Load()
    PHY.Open "Select REF_NO, BAL_QTY, EXP_DATE, SALES_PRICE, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT From RTRXFILE  WHERE ITEM_CODE = '" & FRMSALE.TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY [EXP_DATE]", db, adOpenStatic, adLockReadOnly
    Set grdsub.DataSource = PHY
    grdsub.Columns(0).Caption = "BATCH NO."
    grdsub.Columns(1).Caption = "QTY"
    grdsub.Columns(2).Caption = "EXP DATE"
    grdsub.Columns(3).Caption = "PRICE"
    grdsub.Columns(4).Caption = "TAX"
    grdsub.Columns(5).Caption = "VCH No"
    grdsub.Columns(6).Caption = "Line No"
    grdsub.Columns(7).Caption = "Trx Type"
    
    grdsub.Columns(0).Width = 1400
    grdsub.Columns(1).Width = 900
    grdsub.Columns(2).Width = 1400
    grdsub.Columns(3).Width = 1000
    grdsub.Columns(4).Width = 900
    
    grdsub.Columns(5).Visible = False
    grdsub.Columns(6).Visible = False
    LBLHEAD(0).Caption = grdsub.Columns(6).Text
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PHY.Close
    FRMSALE.Enabled = True
    MDIMAIN.Enabled = True
End Sub

Private Sub grdsub_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            FRMSALE.TXTQTY.Text = grdsub.Columns(1)
            FRMSALE.TXTRATE.Text = grdsub.Columns(3)
            FRMSALE.TXTTAX.Text = grdsub.Columns(4)
            FRMSALE.TXTEXPDATE.Text = grdsub.Columns(2)
            FRMSALE.TXTBATCH.Text = grdsub.Columns(0)
            
            FRMSALE.TXTVCHNO.Text = grdsub.Columns(7)
            FRMSALE.TXTLINENO.Text = grdsub.Columns(8)
            FRMSALE.TXTTRXTYPE.Text = grdsub.Columns(9)
            FRMSALE.TXTUNIT.Text = grdsub.Columns(10)
            
            MDIMAIN.Enabled = True
            FRMSALE.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = False
            FRMSALE.TXTQTY.Enabled = True
            FRMSALE.TXTQTY.SetFocus
            Unload Me
        Case vbKeyEscape
            FRMSALE.TXTQTY.Text = ""
            FRMSALE.TXTVCHNO.Text = ""
            FRMSALE.TXTLINENO.Text = ""
            FRMSALE.TXTTRXTYPE.Text = ""
            FRMSALE.TXTUNIT.Text = ""
            
            MDIMAIN.Enabled = True
            FRMSALE.Enabled = True
            FRMSALE.TXTPRODUCT.Enabled = True
            FRMSALE.TXTQTY.Enabled = False
            FRMSALE.TXTPRODUCT.SetFocus
            Unload Me
    End Select
End Sub

