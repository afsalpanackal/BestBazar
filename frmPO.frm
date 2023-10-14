VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   8955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10560
   Icon            =   "frmPO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8955
   ScaleWidth      =   10560
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
      Height          =   3555
      Left            =   30
      TabIndex        =   38
      Top             =   2220
      Visible         =   0   'False
      Width           =   9555
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   3345
         Left            =   30
         TabIndex        =   39
         Top             =   180
         Width           =   9510
         _ExtentX        =   16775
         _ExtentY        =   5900
         _Version        =   393216
         Rows            =   1
         Cols            =   7
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
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton CmdPrntSum 
      Caption         =   "PRINT SUMMARY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2805
      TabIndex        =   12
      Top             =   8310
      Width           =   1620
   End
   Begin VB.CommandButton CMDPRINTPO 
      Caption         =   "&PRINT PO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1470
      TabIndex        =   11
      Top             =   8310
      Width           =   1245
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&CANCEL"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4620
      TabIndex        =   8
      Top             =   7800
      Width           =   915
   End
   Begin VB.CommandButton cmditemcreate 
      Caption         =   "&Create Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   120
      TabIndex        =   10
      Top             =   8310
      Width           =   1245
   End
   Begin VB.TextBox txtBillNo 
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
      Height          =   315
      Left            =   1065
      TabIndex        =   35
      Top             =   150
      Width           =   885
   End
   Begin VB.CommandButton CMDEXIT 
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
      Height          =   405
      Left            =   5580
      TabIndex        =   9
      Top             =   7800
      Width           =   1065
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   4605
      Left            =   105
      TabIndex        =   23
      Top             =   2385
      Visible         =   0   'False
      Width           =   8460
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   4575
         Left            =   15
         TabIndex        =   24
         Top             =   15
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   8070
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   23
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
   Begin VB.Frame Fram 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frame1"
      Height          =   9000
      Left            =   -120
      TabIndex        =   14
      Top             =   -45
      Width           =   10680
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0FFC0&
         Height          =   1575
         Left            =   150
         TabIndex        =   25
         Top             =   30
         Width           =   10500
         Begin VB.CheckBox chkclosed 
            BackColor       =   &H00004000&
            Caption         =   "Closed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   8370
            TabIndex        =   40
            Top             =   1245
            Width           =   1200
         End
         Begin VB.TextBox TXTLASTBILL 
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
            Height          =   315
            Left            =   7065
            TabIndex        =   33
            Top             =   1230
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.TextBox TXTDATE 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
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
            Height          =   300
            Left            =   2580
            MaxLength       =   10
            TabIndex        =   31
            Top             =   165
            Width           =   1260
         End
         Begin MSDataListLib.DataCombo CMBDISTI 
            Height          =   1020
            Left            =   1035
            TabIndex        =   26
            Top             =   510
            Width           =   5220
            _ExtentX        =   9208
            _ExtentY        =   1799
            _Version        =   393216
            Style           =   1
            ForeColor       =   255
            Text            =   ""
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
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   4830
            TabIndex        =   32
            Top             =   165
            Width           =   1425
            _ExtentX        =   2514
            _ExtentY        =   556
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
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LAST BILL"
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
            Index           =   3
            Left            =   6660
            TabIndex        =   34
            Top             =   1215
            Visible         =   0   'False
            Width           =   1215
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   0
            Left            =   2010
            TabIndex        =   30
            Top             =   180
            Width           =   780
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "PO No."
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
            Index           =   0
            Left            =   120
            TabIndex        =   29
            Top             =   135
            Width           =   870
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "PO Date"
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
            Index           =   8
            Left            =   3990
            TabIndex        =   28
            Top             =   165
            Width           =   855
         End
         Begin VB.Label Label1 
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
            ForeColor       =   &H00FF0000&
            Height          =   300
            Index           =   5
            Left            =   90
            TabIndex        =   27
            Top             =   600
            Width           =   1005
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Height          =   5580
         Left            =   135
         TabIndex        =   15
         Top             =   1515
         Width           =   10515
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
            Left            =   1980
            TabIndex        =   37
            Top             =   960
            Visible         =   0   'False
            Width           =   1080
         End
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   5415
            Left            =   45
            TabIndex        =   13
            Top             =   120
            Width           =   10440
            _ExtentX        =   18415
            _ExtentY        =   9551
            _Version        =   393216
            Rows            =   1
            Cols            =   7
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00C0FFC0&
         Height          =   1935
         Left            =   135
         TabIndex        =   16
         Top             =   7020
         Width           =   10515
         Begin VB.TextBox TXTUOM 
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
            Height          =   300
            Left            =   7020
            MaxLength       =   7
            TabIndex        =   3
            Top             =   450
            Width           =   990
         End
         Begin VB.TextBox TXTRCVDQTY 
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
            Height          =   300
            Left            =   5535
            TabIndex        =   36
            Top             =   1350
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.CommandButton CMDMODIFY 
            Caption         =   "&Modify Line"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   120
            TabIndex        =   4
            Top             =   825
            Width           =   1155
         End
         Begin VB.TextBox TXTSLNO 
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
            Height          =   300
            Left            =   120
            TabIndex        =   0
            Top             =   450
            Width           =   645
         End
         Begin VB.TextBox TXTPRODUCT 
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
            Height          =   300
            Left            =   780
            TabIndex        =   1
            Top             =   450
            Width           =   5355
         End
         Begin VB.TextBox TXTQTY 
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
            Height          =   300
            Left            =   6150
            MaxLength       =   7
            TabIndex        =   2
            Top             =   450
            Width           =   855
         End
         Begin VB.CommandButton cmdadd 
            Caption         =   "&ADD"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2550
            TabIndex        =   6
            Top             =   825
            Width           =   1020
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "&Delete Line"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1365
            TabIndex        =   5
            Top             =   825
            Width           =   1155
         End
         Begin VB.TextBox TXTITEMCODE 
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
            Height          =   300
            Left            =   5850
            TabIndex        =   17
            Top             =   765
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&SAVE"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   3600
            TabIndex        =   7
            Top             =   825
            Width           =   960
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "UOM"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   1
            Left            =   7020
            TabIndex        =   41
            Top             =   195
            Width           =   990
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SL No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   22
            Top             =   195
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Product Name"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   270
            Index           =   9
            Left            =   780
            TabIndex        =   21
            Top             =   195
            Width           =   5355
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Qty"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Index           =   10
            Left            =   6150
            TabIndex        =   20
            Top             =   195
            Width           =   855
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "ITEM CODE."
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
            Height          =   255
            Index           =   15
            Left            =   4620
            TabIndex        =   19
            Top             =   780
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "LINE NO."
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
            Height          =   255
            Index           =   18
            Left            =   4725
            TabIndex        =   18
            Top             =   1770
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "frmPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT As Boolean
Dim M_ADD As Boolean

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBDISTI.text = "" Then Exit Sub
            If IsNull(CMBDISTI.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "PURCHASE ORDER"
                CMBDISTI.SetFocus
                Exit Sub
            End If
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            If M_ADD = False Then
                FRMECONTROLS.Enabled = False
                FRMEMASTER.Enabled = False
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
            End If
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMDADD_Click()
    Dim i As Long
    
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = Trim(TXTITEMCODE.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(TXTPRODUCT.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Val(TXTQTY.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = Val(TXTRCVDQTY.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 6) = Trim(TXTUOM.text)
    Select Case Val(TXTQTY.text) - Val(TXTRCVDQTY.text)
        Case 0
            grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "S" 'Same
        Case Is > 0
            grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "L" 'Less
        Case Is < 0
            grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "M" 'More
    End Select
    
    TXTSLNO.text = grdsales.rows
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    TXTRCVDQTY.text = ""
    TXTQTY.text = ""
    TXTUOM.text = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    cmdRefresh.Enabled = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    txtBillNo.Enabled = False
    
    'If grdsales.Rows >= 18 Then grdsales.TopRow = grdsales.Rows - 1
        
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTQTY.Enabled = True
            TXTUOM.Enabled = True
            TXTQTY.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdcancel_Click()
    If M_ADD = True Then
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo, "PURCHASE ORDER...") = vbNo Then Exit Sub
    End If
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    CMBDISTI.text = ""
    TXTINVDATE.text = "  /  /    "
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    TXTQTY.text = ""
    TXTUOM.text = ""
    TXTRCVDQTY = ""
    chkclosed.Value = 0
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    'grdsales.TextArray(0) = "SL"
    'grdsales.TextArray(1) = "ITEM CODE"
    'grdsales.TextArray(2) = "ITEM NAME"
    'grdsales.TextArray(3) = "QTY"
    'grdsales.TextArray(5) = "MRP"
    'grdsales.TextArray(6) = "RATE"
    'grdsales.TextArray(7) = "PTR"
    'grdsales.TextArray(8) = "COST"
    'grdsales.TextArray(9) = "Serial No"
    'grdsales.TextArray(11) = "SUB TOTAL"
    
    For i = Val(TXTSLNO.text) To grdsales.rows - 2
        grdsales.TextMatrix(i, 0) = i
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(i, 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(i, 6) = grdsales.TextMatrix(i + 1, 6)
    Next i
    grdsales.rows = grdsales.rows - 1
    TXTSLNO.text = Val(grdsales.rows)
    TXTPRODUCT.text = ""
    TXTITEMCODE.text = ""
    TXTQTY.text = ""
    TXTUOM.text = ""
    TXTRCVDQTY = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
    If grdsales.rows = 1 Then
        CMDEXIT.Enabled = True
    End If
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub cmditemcreate_Click()
    frmitemmaster.Show
    frmitemmaster.SetFocus
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTQTY.Enabled = True
    TXTUOM.Enabled = True
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.text = grdsales.rows
            TXTPRODUCT.text = ""
            TXTQTY.text = ""
            TXTUOM.text = ""
            TXTRCVDQTY.text = ""
            TXTITEMCODE.text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTUOM.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub CMDPRINTPO_Click()
    Dim i As Long
    On Error GoTo ErrHand
    If grdsales.rows = 1 Then
        MsgBox "Please Select Purchase Order No.", vbOKOnly, "PURCHASE ORDER..."
        Exit Sub
    End If
    If M_ADD = True Then
        Dim rstMaxNo As ADODB.Recordset
    
        If IsNull(CMBDISTI.SelectedItem) Then
            MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        If Not IsDate(TXTINVDATE.text) Then
            MsgBox "Enter Purchase Order Date", vbOKOnly, "EzBiz"
            Exit Sub
        End If
        Call appendpurchase
        M_ADD = False
    End If
    Sleep (300)
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompAddress4, CompTin, CompCST As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        CompAddress3 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO)
        CompAddress4 = IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!EMAIL_ADD)
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    ReportNameVar = Rptpath & "RPTPurchase"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Report.RecordSelectionFormula = "({POSUB.VCH_NO}= " & Val(txtBillNo.text) & " )"
    Set CRXFormulaFields = Report.FormulaFields
    
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
        If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.text = "'" & CompAddress4 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
    Next
    frmreport.Caption = "PURCHASE ORDER"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub CmdPrntSum_Click()
    Dim i As Long
    If grdsales.rows = 1 Then
        MsgBox "Please Select Purchase Order No.", vbOKOnly, "PURCHASE ORDER..."
        Exit Sub
    End If
    If M_ADD = True Then
        Dim rstMaxNo As ADODB.Recordset
    
        If IsNull(CMBDISTI.SelectedItem) Then
            MsgBox "Select Supplier From List", vbOKOnly, "PURCHASE ORDER"
            Exit Sub
        End If
        If Not IsDate(TXTINVDATE.text) Then
            MsgBox "Enter Purchase Order Date", vbOKOnly, "PURCHASE ORDER"
            Exit Sub
        End If
        Call appendpurchase
        M_ADD = False
    End If
    Sleep (300)
    
    On Error GoTo ErrHand
    
    ReportNameVar = Rptpath & "RPTPurchaseSum"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "( {POSUB.QTY}-{POSUB.RCVD_QTY}<>0 and {POSUB.VCH_NO}= " & Val(txtBillNo.text) & " )"
    ''Report.RecordSelectionFormula = "( {POSUB.VCH_NO}= " & Val(txtBillNo.Text) & " )"
    Set CRXFormulaFields = Report.FormulaFields
    
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
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "PURCHASE ORDER VARIATIONS"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
    Dim rstMaxNo As ADODB.Recordset
    
    If IsNull(CMBDISTI.SelectedItem) Then
        MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Purchase Order Date", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Call appendpurchase
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From POMAST", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    chkclosed.Value = 0
    grdsales.rows = 1
    TXTSLNO.text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    CMBDISTI.text = ""
    TXTINVDATE.text = "  /  /    "
    TXTSLNO.text = ""
    TXTITEMCODE.text = ""
    TXTPRODUCT.text = ""
    TXTQTY.text = ""
    TXTUOM.text = ""
    TXTRCVDQTY.text = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "PURCHASE ORDER"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    On Error GoTo ErrHand
    txtBillNo.SetFocus
    Exit Sub
ErrHand:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From POMAST", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    Call fillcombo
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 1300
    grdsales.ColWidth(2) = 5000
    grdsales.ColWidth(3) = 1100
    grdsales.ColWidth(4) = 1200
    grdsales.ColWidth(6) = 1000
    grdsales.ColWidth(5) = 0
    
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(1) = 1
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 7
    grdsales.ColAlignment(4) = 7
    grdsales.ColAlignment(5) = 7
    grdsales.ColAlignment(6) = 7
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "RCVD QTY"
    grdsales.TextArray(5) = "FLAG"
    grdsales.TextArray(6) = "UOM"

    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Qty"
    GRDBILL.TextMatrix(0, 3) = "Rate"
    GRDBILL.TextMatrix(0, 4) = "Total"
    GRDBILL.TextMatrix(0, 5) = "Bill No"
    GRDBILL.TextMatrix(0, 6) = "Date"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 3500
    GRDBILL.ColWidth(2) = 1000
    GRDBILL.ColWidth(3) = 1000
    GRDBILL.ColWidth(4) = 1200
    GRDBILL.ColWidth(5) = 1000
    GRDBILL.ColWidth(6) = 1200
    
    GRDBILL.ColAlignment(0) = 4
    GRDBILL.ColAlignment(2) = 4
    GRDBILL.ColAlignment(3) = 4
    GRDBILL.ColAlignment(4) = 4
    GRDBILL.ColAlignment(5) = 4
    GRDBILL.ColAlignment(6) = 4
    
    PHYFLAG = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTUOM.Enabled = False
    TXTINVDATE.text = Date
    TXTDATE.text = Date
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTSLNO.text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    M_ADD = False
    'Me.Width = 7665
    'Me.Height = 9435
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ACT_FLAG = False Then ACT_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub


Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn

            TXTPRODUCT.text = grdtmp.Columns(1)
            TXTITEMCODE.text = grdtmp.Columns(0)
            TXTUOM.text = grdtmp.Columns(3)
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTUOM.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.text = ""
            TXTUOM.text = ""
            TXTRCVDQTY.text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.rows = 1
            i = 0
            grdsales.rows = 1
            Dim RSTITEMMAST As ADODB.Recordset
            Dim RCVDQTY As Double
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From POSUB WHERE VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rstTRXMAST.EOF
                RCVDQTY = 0
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * from RTRXFILE WHERE PO_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE = '" & rstTRXMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
                Do Until RSTITEMMAST.EOF
                    RCVDQTY = RCVDQTY + IIf(IsNull(RSTITEMMAST!QTY), 0, RSTITEMMAST!QTY)
                    RSTITEMMAST.MoveNext
                Loop
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                rstTRXMAST!RCVD_QTY = RCVDQTY
                rstTRXMAST.Update
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From POSUB WHERE VCH_NO = " & Val(txtBillNo.text) & " ORDER BY LINE_NO", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 3) = Val(rstTRXMAST!QTY)
                grdsales.TextMatrix(i, 6) = IIf(IsNull(rstTRXMAST!UNIT), "Nos", rstTRXMAST!UNIT)
                grdsales.TextMatrix(i, 4) = IIf(IsNull(rstTRXMAST!RCVD_QTY), 0, Val(rstTRXMAST!RCVD_QTY))
                Select Case Val(grdsales.TextMatrix(i, 3)) - Val(grdsales.TextMatrix(i, 4))
                    Case 0
                        grdsales.TextMatrix(i, 5) = "S" 'Same
                    Case Is > 0
                        grdsales.TextMatrix(i, 5) = "L" 'Less
                    Case Is < 0
                        grdsales.TextMatrix(i, 5) = "M" 'More
                End Select
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From POMAST WHERE VCH_NO = " & Val(txtBillNo.text) & "", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                CMBDISTI.text = rstTRXMAST!ACT_NAME
                If IsNull(rstTRXMAST!Status) Or rstTRXMAST!Status = "N" Then
                    chkclosed.Value = 0
                Else
                    chkclosed.Value = 1
                End If
                cmdRefresh.Enabled = True
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            TXTSLNO.text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.text) < Val(TXTLASTBILL.text)) Then
                'FRMEMASTER.Enabled = False
                cmdcancel.SetFocus
            Else
                CMBDISTI.SetFocus
            End If
    End Select
    
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
    If Val(txtBillNo.text) = 0 Or Val(txtBillNo.text) > Val(TXTLASTBILL.text) Then txtBillNo.text = TXTLASTBILL.text
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.text = "  /  /    " Then
                TXTINVDATE.text = Format(Date, "DD/MM/YYYY")
                FRMECONTROLS.Enabled = True
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.text) Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.text = Format(TXTINVDATE.text, "DD/MM/YYYY")
                FRMECONTROLS.Enabled = True
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            End If
        Case vbKeyEscape
            CMBDISTI.SetFocus
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

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTPRODUCT.text) = "" Then Exit Sub
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, PACK_TYPE From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTPRODUCT.text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.text = grdtmp.Columns(0)
                TXTPRODUCT.text = grdtmp.Columns(1)
                TXTUOM.text = grdtmp.Columns(3)
                For i = 1 To grdsales.rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i
                
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTUOM.Enabled = True
                    TXTQTY.SetFocus
                    Exit Sub
                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 5500
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1300
                grdtmp.Columns(3).Caption = "UOM"
                grdtmp.Columns(3).Width = 1000
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            'TXTQTY.Enabled = False
            'TXTRATE.Enabled = False
            'txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(Chr(KeyAscii))
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 Then Exit Sub
            TXTQTY.Enabled = False
            TXTUOM.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTUOM.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.text = Format(TXTQTY.text, ".00")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.text) = 0 Then
                TXTSLNO.text = grdsales.rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.text) >= grdsales.rows Then
                TXTSLNO.text = grdsales.rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            
            If Val(TXTSLNO.text) < grdsales.rows Then
                TXTSLNO.text = grdsales.TextMatrix(Val(TXTSLNO.text), 0)
                TXTITEMCODE.text = grdsales.TextMatrix(Val(TXTSLNO.text), 1)
                TXTPRODUCT.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                TXTQTY.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3))
                TXTUOM.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 6))
                TXTRCVDQTY.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 4))
                Select Case Val(TXTQTY.text) - Val(TXTRCVDQTY.text)
                    Case 0
                        grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "S" 'Same
                    Case Is > 0
                        grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "L" 'Less
                    Case Is < 0
                        grdsales.TextMatrix(Val(TXTSLNO.text), 5) = "M" 'More
                End Select
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTUOM.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTUOM.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.text = Val(grdsales.rows)
                TXTPRODUCT.text = ""
                TXTITEMCODE.text = ""
                TXTQTY.text = ""
                TXTUOM.text = ""
                TXTRCVDQTY.text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.rows > 1 Then
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            End If
            If M_ADD = False Then
                FRMECONTROLS.Enabled = False
                FRMEMASTER.Enabled = False
                cmdRefresh.Enabled = False
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
                Exit Sub
            End If
            
    End Select
End Sub

Private Sub TXTSLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case vbKeyTab
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub fillcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Public Sub appendpurchase()
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    db.Execute "delete From POMAST WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From POSUB WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From POMAST", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = CMBDISTI.BoundText
    RSTTRXFILE!ACT_NAME = Trim(CMBDISTI.text)
    If chkclosed.Value = 1 Then
        RSTTRXFILE!Status = "Y"
    Else
        RSTTRXFILE!Status = "N"
    End If
    RSTTRXFILE!CREATE_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!MODIFY_DATE = Format(TXTDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    For i = 1 To grdsales.rows - 1
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * from POSUB", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTRTRXFILE!VCH_DATE = Format(Trim(TXTINVDATE.text), "dd/mm/yyyy")
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(i, 0))
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(i, 1))
        RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTRTRXFILE!RCVD_QTY = Val(grdsales.TextMatrix(i, 4))
        RSTRTRXFILE!UNIT = Trim(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTRTRXFILE!C_USER_ID = "SM"
        RSTRTRXFILE!M_USER_ID = Left(CMBDISTI.BoundText, 8)
        RSTRTRXFILE.Update
        
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing

    Next i
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdsales.Col
                Case 3, 4 ' Bal QTY
                    grdsales.TextMatrix(grdsales.Row, grdsales.Col) = Val(TXTsample.text)
                    Select Case Val(grdsales.TextMatrix(grdsales.Row, 3)) - Val(grdsales.TextMatrix(grdsales.Row, 4))
                        Case 0
                            grdsales.TextMatrix(grdsales.Row, 5) = "S" 'Same
                        Case Is > 0
                            grdsales.TextMatrix(grdsales.Row, 5) = "L" 'Less
                        Case Is < 0
                            grdsales.TextMatrix(grdsales.Row, 5) = "M" 'More
                    End Select
                    Call appendpurchase
                    M_ADD = False
                    grdsales.Enabled = True
                    TXTsample.Visible = False
                    grdsales.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdsales.SetFocus
    End Select
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdsales.Col
        Case 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub grdsales_Scroll()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub grdsales_Click()
    TXTsample.Visible = False
    grdsales.SetFocus
End Sub

Private Sub grdsales_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            
            If grdsales.rows = 1 Then Exit Sub
            GRDBILL.rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT * from RTRXFILE WHERE PO_NO = " & Val(txtBillNo.text) & " AND ITEM_CODE = '" & grdsales.TextMatrix(grdsales.Row, 1) & "' ", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.rows = GRDBILL.rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!QTY, "0.00")
                'GRDBILL.TextMatrix(i, 3) = Val(RSTTRXFILE!LINE_DISC)
                'GRDBILL.TextMatrix(i, 4) = Val(RSTTRXFILE!SALES_TAX)
                'GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!QTY
                GRDBILL.TextMatrix(i, 4) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                GRDBILL.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!VCH_NO), "", RSTTRXFILE!VCH_NO)
                GRDBILL.TextMatrix(i, 6) = IIf(IsNull(RSTTRXFILE!VCH_DATE), "", RSTTRXFILE!VCH_DATE)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing

            FRMEBILL.Visible = True
            GRDBILL.SetFocus
    End Select
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEBILL.Visible = False
            grdsales.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        FRMEBILL.Visible = False
        grdsales.SetFocus
    End If
End Sub

Private Sub TXTUOM_GotFocus()
    TXTUOM.SelStart = 0
    TXTUOM.SelLength = Len(TXTUOM.text)
End Sub

Private Sub TXTUOM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTUOM.text) = "" Then Exit Sub
            TXTQTY.Enabled = False
            TXTUOM.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.Enabled = True
            TXTUOM.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUOM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTUOM_LostFocus()
    TXTUOM.text = Format(TXTUOM.text, ".00")
End Sub

