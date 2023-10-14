VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMCRDTNOTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDIT NOTE"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   ControlBox      =   0   'False
   Icon            =   "FrmSTOCK.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   9615
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2865
      Left            =   1995
      TabIndex        =   40
      Top             =   2040
      Visible         =   0   'False
      Width           =   5010
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   2490
         Left            =   165
         TabIndex        =   41
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   4392
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
   End
   Begin VB.Frame Fram 
      Caption         =   "Frame1"
      Height          =   7830
      Left            =   0
      TabIndex        =   16
      Top             =   -45
      Width           =   9480
      Begin VB.Frame Frame1 
         Height          =   690
         Left            =   45
         TabIndex        =   17
         Top             =   45
         Width           =   9375
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
            Height          =   345
            Left            =   975
            TabIndex        =   13
            Top             =   210
            Width           =   885
         End
         Begin VB.TextBox TXTREMARKS 
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
            Height          =   345
            Left            =   6060
            TabIndex        =   14
            Top             =   225
            Width           =   2115
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
            Height          =   300
            Index           =   0
            Left            =   105
            TabIndex        =   22
            Top             =   240
            Width           =   870
         End
         Begin VB.Label Label1 
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
            Height          =   300
            Index           =   1
            Left            =   1920
            TabIndex        =   21
            Top             =   240
            Width           =   645
         End
         Begin VB.Label LBLDATE 
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
            Height          =   360
            Left            =   2490
            TabIndex        =   20
            Top             =   225
            Width           =   1215
         End
         Begin VB.Label LBLTIME 
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
            Height          =   360
            Left            =   3735
            TabIndex        =   19
            Top             =   225
            Width           =   1110
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "REMARKS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   3
            Left            =   5025
            TabIndex        =   18
            Top             =   255
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         Height          =   5730
         Left            =   60
         TabIndex        =   23
         Top             =   600
         Width           =   9360
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4950
            Left            =   90
            TabIndex        =   15
            Top             =   270
            Width           =   9045
            _ExtentX        =   15954
            _ExtentY        =   8731
            _Version        =   393216
            Rows            =   1
            Cols            =   12
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
         End
      End
      Begin VB.Frame Frame4 
         Height          =   1575
         Left            =   60
         TabIndex        =   24
         Top             =   6240
         Width           =   9390
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
            Height          =   480
            Left            =   2895
            TabIndex        =   8
            Top             =   795
            Width           =   1125
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
            Left            =   165
            TabIndex        =   0
            Top             =   450
            Width           =   570
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
            Left            =   750
            TabIndex        =   1
            Top             =   450
            Width           =   3420
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
            Left            =   4950
            MaxLength       =   7
            TabIndex        =   2
            Top             =   450
            Width           =   675
         End
         Begin VB.TextBox TXTRATE 
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
            Left            =   5655
            MaxLength       =   6
            TabIndex        =   3
            Top             =   450
            Width           =   630
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
            Height          =   465
            Left            =   7725
            TabIndex        =   12
            Top             =   810
            Width           =   1125
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
            Height          =   465
            Left            =   5340
            TabIndex        =   10
            Top             =   810
            Width           =   1125
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
            Height          =   480
            Left            =   4095
            TabIndex        =   9
            Top             =   795
            Width           =   1125
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
            Left            =   1275
            TabIndex        =   28
            Top             =   825
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox txtBatch 
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
            Height          =   285
            Left            =   7425
            MaxLength       =   15
            TabIndex        =   6
            Top             =   435
            Width           =   930
         End
         Begin VB.TextBox TXTVCHNO 
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
            Left            =   1260
            TabIndex        =   27
            Top             =   1200
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TXTLINENO 
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
            Left            =   3120
            TabIndex        =   26
            Top             =   1200
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.TextBox TXTUNIT 
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
            Left            =   4200
            TabIndex        =   25
            Top             =   450
            Width           =   720
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "&Refresh"
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
            Height          =   465
            Left            =   6540
            TabIndex        =   11
            Top             =   810
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   285
            Left            =   6315
            TabIndex        =   4
            Top             =   450
            Visible         =   0   'False
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   5
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##"
            PromptChar      =   " "
         End
         Begin VB.TextBox txtexpdate 
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
            Left            =   6300
            MaxLength       =   10
            TabIndex        =   5
            Top             =   435
            Width           =   1095
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
            Height          =   225
            Index           =   8
            Left            =   165
            TabIndex        =   39
            Top             =   225
            Width           =   570
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
            Height          =   240
            Index           =   9
            Left            =   750
            TabIndex        =   38
            Top             =   225
            Width           =   3420
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
            Height          =   225
            Index           =   10
            Left            =   4950
            TabIndex        =   37
            Top             =   225
            Width           =   675
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Rate"
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
            Height          =   225
            Index           =   11
            Left            =   5655
            TabIndex        =   36
            Top             =   225
            Width           =   630
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Sub Total"
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
            Height          =   225
            Index           =   14
            Left            =   8385
            TabIndex        =   35
            Top             =   210
            Width           =   930
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
            Left            =   135
            TabIndex        =   34
            Top             =   840
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Exp Date"
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
            Height          =   225
            Index           =   16
            Left            =   6300
            TabIndex        =   33
            Top             =   225
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Batch"
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
            Height          =   225
            Index           =   7
            Left            =   7425
            TabIndex        =   32
            Top             =   225
            Width           =   930
         End
         Begin VB.Label LBLSUBTOTAL 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   300
            Left            =   8385
            TabIndex        =   7
            Top             =   420
            Width           =   930
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "VCH NO."
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
            Index           =   17
            Left            =   120
            TabIndex        =   31
            Top             =   1215
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "UNIT"
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
            Index           =   20
            Left            =   4200
            TabIndex        =   30
            Top             =   225
            Width           =   720
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
            Left            =   1980
            TabIndex        =   29
            Top             =   1215
            Visible         =   0   'False
            Width           =   1080
         End
      End
   End
End
Attribute VB_Name = "FRMCRDTNOTE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT As Boolean

Private Sub CMDADD_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Integer

    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTUNIT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TXTRATE.Text) * Val(TXTUNIT.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Trim(txtexpdate.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = txtBillNo.Text
    grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Val(TXTSLNO.Text)
    

    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTUNIT.Text = ""
    
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    txtBatch.Text = ""
    txtexpdate.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    txtBillNo.Enabled = False
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Integer
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    
    For i = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = i
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(i + 1, 4)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = grdsales.TextMatrix(i + 1, 5)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = grdsales.TextMatrix(i + 1, 6)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = grdsales.TextMatrix(i + 1, 7)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = grdsales.TextMatrix(i + 1, 8)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = grdsales.TextMatrix(i + 1, 9)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = grdsales.TextMatrix(i + 1, 10)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = i
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTUNIT.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    txtexpdate.Text = ""
    txtBatch.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    
    If grdsales.Rows = 1 Then
        CMDEXIT.Enabled = True
    End If
End Sub

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTUNIT.Text = ""
            LBLSUBTOTAL.Caption = ""
            txtexpdate.Text = ""
            txtBatch.Text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            txtexpdate.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    
    Dim RSTSTOCK As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RSTRTRXFILE As ADODB.Recordset
    Dim rstMaxNo As ADODB.Recordset
    
    Dim M_DATA As Integer
    Dim i As Integer
    
    On Error GoTo ErrHand
    db.Execute "delete * From RTRXFILE WHERE TRX_TYPE='CN' AND VCH_NO = " & Val(txtBillNo.Text) & ""
    
    For i = 1 To grdsales.Rows - 1
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * from [RTRXFILE]", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTRTRXFILE.AddNew
        RSTRTRXFILE!TRX_TYPE = "CN"
        RSTRTRXFILE!VCH_NO = Val(grdsales.TextMatrix(i, 10))
        RSTRTRXFILE!VCH_DATE = Format(Trim(LBLDATE.Caption), "dd/mm/yyyy")
        RSTRTRXFILE!LINE_NO = Val(grdsales.TextMatrix(i, 11))
        RSTRTRXFILE!CATEGORY = "MEDICINE"
        RSTRTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(i, 1))
        RSTRTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTRTRXFILE!QTY = Val(grdsales.TextMatrix(i, 3))
        RSTRTRXFILE!ITEM_COST = 0
        RSTRTRXFILE!MRP = Val(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!PTR = Val(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(i, 6))
        RSTRTRXFILE!SALES_TAX = 0
        RSTRTRXFILE!UNIT = Val(grdsales.TextMatrix(i, 4))
        RSTRTRXFILE!VCH_DESC = "C/Note from " & Trim(TXTREMARKS.Text)
        RSTRTRXFILE!REF_NO = Trim(grdsales.TextMatrix(i, 7))
        RSTRTRXFILE!ISSUE_QTY = 0
        RSTRTRXFILE!CST = 0
        RSTRTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(i, 3)) '* Val(grdsales.TextMatrix(i, 4))
        RSTRTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(i, 9))
        RSTRTRXFILE!LINE_DISC = 0
        RSTRTRXFILE!SCHEME = 0
        RSTRTRXFILE!EXP_DATE = Format(Trim(grdsales.TextMatrix(i, 8)), "dd/mm/yyyy")
        RSTRTRXFILE!FREE_QTY = 0
        RSTRTRXFILE!CREATE_DATE = Date
        RSTRTRXFILE!C_USER_ID = "SM"
        RSTRTRXFILE!M_USER_ID = "311000"
        RSTRTRXFILE!CHECK_FLAG = ""
        RSTRTRXFILE!PINV = ""
        RSTRTRXFILE.Update
        
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing

    Next i
    
    E_DATA = 0
    For i = 1 To grdsales.Rows - 1
    
        E_DATA = 0
        Set RSTSTOCK = New ADODB.Recordset
        RSTSTOCK.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockReadOnly
        Do Until RSTSTOCK.EOF
            E_DATA = E_DATA + RSTSTOCK!BAL_QTY
            RSTSTOCK.MoveNext
        Loop
        RSTSTOCK.Close
        Set RSTSTOCK = Nothing
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(i, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTITEMMAST
            If Not (.EOF And .BOF) Then
                !OPEN_QTY = E_DATA
                !OPEN_VAL = 0
                !RCPT_QTY = 0
                !RCPT_VAL = 0
                !ISSUE_QTY = 0
                !ISSUE_VAL = 0
                !CLOSE_QTY = E_DATA
                !CLOSE_VAL = 0
                !DAM_QTY = 0
                !DAM_VAL = 0
                RSTITEMMAST.Update
            End If
        End With
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing

    Next i
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(Val(VCH_NO)) From RTRXFILE WHERE TRX_TYPE = 'CN'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    CMDREFRESH.Enabled = False
    txtBillNo.Enabled = True
    CMDEXIT.Enabled = True
    CMDEXIT.Enabled = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From RTRXFILE WHERE TRX_TYPE = 'CN'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing

    
    LBLDATE.Caption = Date
    LBLTIME.Caption = Time
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2700
    grdsales.ColWidth(3) = 500
    grdsales.ColWidth(4) = 500
    grdsales.ColWidth(5) = 800
    grdsales.ColWidth(6) = 800
    grdsales.ColWidth(7) = 1000
    grdsales.ColWidth(8) = 1000
    grdsales.ColWidth(9) = 1100
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "UNIT"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "BATCH"
    grdsales.TextArray(8) = "EXPIRY"
    grdsales.TextArray(9) = "SUB TOTAL"
    grdsales.TextArray(10) = "Vch No"
    grdsales.TextArray(11) = "Line No"
        

    PHYFLAG = True
    TMPFLAG = True
    TXTPRODUCT.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    txtexpdate.Enabled = False
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTREMARKS.Text = "SALES RETURN"
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    CLOSEALL = 1
    
    Me.Width = 9800
    Me.Height = 8300
    Me.Left = 0
    Me.Top = 0
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn

            TXTPRODUCT.Text = grdtmp.Columns(1)
            TXTITEMCODE.Text = grdtmp.Columns(0)
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select DISTINCT [ITEM_CODE],[UNIT] From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                TXTUNIT.Text = RSTRXFILE!UNIT
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTUNIT.Enabled = True
            TXTUNIT.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTUNIT.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTUNIT.Text = ""
            Fram.Enabled = True
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(txtBatch.Text) = "" Then Exit Sub
            txtBatch.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            txtexpdate.Enabled = True
            txtBatch.Enabled = False
            txtexpdate.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTTRXMAST As ADODB.Recordset
Dim i As Integer

    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.Rows = 1
           
            i = 0
            grdsales.Rows = 1
            Set RSTTRXMAST = New ADODB.Recordset
            RSTTRXMAST.Open "Select * From RTRXFILE WHERE TRX_TYPE='CN' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY [LINE_NO]", db, adOpenStatic, adLockReadOnly
            Do Until RSTTRXMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = RSTTRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 2) = RSTTRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 3) = RSTTRXMAST!QTY
                grdsales.TextMatrix(i, 4) = RSTTRXMAST!UNIT
                grdsales.TextMatrix(i, 5) = Format(RSTTRXMAST!MRP * Val(RSTTRXMAST!UNIT), ".000")
                grdsales.TextMatrix(i, 6) = Format(RSTTRXMAST!MRP, ".000")
                grdsales.TextMatrix(i, 7) = RSTTRXMAST!REF_NO
                grdsales.TextMatrix(i, 8) = Format(RSTTRXMAST!EXP_DATE, "DD/MM/YYYY")
                grdsales.TextMatrix(i, 9) = Format(RSTTRXMAST!TRX_TOTAL, ".000")
                
                grdsales.TextMatrix(i, 10) = RSTTRXMAST!VCH_NO
                grdsales.TextMatrix(i, 11) = RSTTRXMAST!LINE_NO
                TXTREMARKS.Text = Mid(RSTTRXMAST!VCH_DESC, 12)
                RSTTRXMAST.MoveNext
            Loop
            RSTTRXMAST.Close
            Set RSTTRXMAST = Nothing
            
            txtBillNo.Enabled = False
            TXTREMARKS.SetFocus
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = True
    End Select
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTBILLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtBillNo_LostFocus()
Dim TRXMAST As ADODB.Recordset
Dim i As Double

i = 1
Set TRXMAST = New ADODB.Recordset
TRXMAST.Open "Select MAX(Val(VCH_NO)) From RTRXFILE WHERE TRX_TYPE = 'CN'", db, adOpenStatic, adLockReadOnly
If Not (TRXMAST.EOF And TRXMAST.BOF) Then
    i = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    If Val(txtBillNo.Text) > i Or Val(txtBillNo.Text) = 0 Then
        txtBillNo.Text = i
        Exit Sub
    End If
End If
TRXMAST.Close
Set TRXMAST = Nothing

End Sub

Private Sub TXTEXPDATE_GotFocus()
    txtexpdate.SelStart = 0
    txtexpdate.SelLength = Len(txtexpdate.Text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(txtexpdate.Text) Then Exit Sub
            txtBatch.Enabled = True
            TXTEXPIRY.Visible = False
            txtexpdate.Enabled = False
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTRATE.Enabled = True
            txtexpdate.Enabled = False
            TXTEXPIRY.Visible = False
            TXTRATE.SetFocus
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKey0 To vbKey9, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub


Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTRXFILE As ADODB.Recordset
Dim i As Integer
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME], CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TXTPRODUCT.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenStatic, adLockReadOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i

                Set RSTRXFILE = New ADODB.Recordset
                RSTRXFILE.Open "Select DISTINCT [ITEM_CODE],[UNIT] From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    TXTUNIT.Text = RSTRXFILE!UNIT
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                If PHY.RecordCount = 1 Then
                    TXTPRODUCT.Enabled = False
                    TXTUNIT.Enabled = True
                    TXTUNIT.SetFocus
                    Exit Sub
                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 3000
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1100
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
        
            TXTQTY.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
         Case vbKeyEscape
            TXTUNIT.Enabled = True
            TXTQTY.Enabled = False
            TXTUNIT.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)), ".000")
End Sub

Private Sub TXTRATE_GotFocus()
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = False
            TXTEXPIRY.Visible = True
            TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTQTY.Enabled = True
            TXTRATE.Enabled = False
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRATE_LostFocus()
    TXTRATE.Text = Format(Val(TXTRATE.Text) / Val(TXTUNIT.Text), ".000")
    'TXTRATE.Text = Format(TXTRATE.Text, ".000")
    LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)), ".000")
End Sub

Private Sub TXTREMARKS_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub TXTREMARKS_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = grdsales.Rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.Rows Then
                TXTSLNO.Text = grdsales.Rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.Rows Then
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                TXTITEMCODE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                TXTQTY.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 3)
                TXTUNIT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                txtexpdate.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                TXTEXPIRY.Text = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 8), "mm/yy")
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                
                TXTSLNO.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                txtexpdate.Enabled = False
                txtBatch.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = False
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTUNIT.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                LBLSUBTOTAL.Caption = ""
                txtexpdate.Text = ""
                txtBatch.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                CMDREFRESH.Enabled = True
                CMDREFRESH.SetFocus
            End If
    End Select
End Sub

Private Sub TXTSLNO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case vbKeyTab
            KeyAscii = 0
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then Exit Sub
            TXTEXPIRY.Visible = False
            txtexpdate.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            TXTRATE.Enabled = True
            txtexpdate.Enabled = False
            TXTRATE.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then
        txtexpdate.Text = ""
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then
        txtexpdate.Text = ""
        Exit Sub
    End If
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then
        txtexpdate.Text = ""
        Exit Sub
    End If
    
    M = Val(Mid(TXTEXPIRY.Text, 1, 2))
    Y = Val(Right(TXTEXPIRY.Text, 2))
    Y = 2000 + Y
    M_DATE = "01" & "/" & M & "/" & Y
    D = LastDayOfMonth(M_DATE)
    M_DATE = D & "/" & M & "/" & Y
    txtexpdate.Text = M_DATE
    txtexpdate.SelStart = 0
    txtexpdate.SelLength = Len(txtBatch.Text)
    TXTEXPIRY.Visible = False
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            
            TXTUNIT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTUNIT.Text = ""
            TXTPRODUCT.Enabled = True
            TXTUNIT.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
