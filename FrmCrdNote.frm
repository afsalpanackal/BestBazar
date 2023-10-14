VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMCRDTNOTE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SALES RETURN"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   14310
   Begin VB.Frame fRMEPRERATE 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3420
      Left            =   3450
      TabIndex        =   66
      Top             =   2250
      Visible         =   0   'False
      Width           =   10320
      Begin MSDataGridLib.DataGrid GRDPRERATE 
         Height          =   2910
         Left            =   30
         TabIndex        =   67
         Top             =   480
         Width           =   10260
         _ExtentX        =   18098
         _ExtentY        =   5133
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
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
         Caption         =   " SOLD RATES FOR THE ITEM "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   360
         Index           =   1
         Left            =   30
         TabIndex        =   69
         Top             =   105
         Width           =   3360
      End
      Begin VB.Label LBLHEAD 
         BackColor       =   &H00000000&
         Caption         =   "MEDICINE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   360
         Index           =   2
         Left            =   3390
         TabIndex        =   68
         Top             =   105
         Width           =   6900
      End
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   2070
      TabIndex        =   65
      Top             =   0
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   0
      ScaleHeight     =   240
      ScaleWidth      =   2070
      TabIndex        =   63
      Top             =   0
      Visible         =   0   'False
      Width           =   2070
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         FillStyle       =   0  'Solid
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   0
         ScaleHeight     =   240
         ScaleWidth      =   2070
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   2070
      End
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
      Left            =   8745
      TabIndex        =   17
      Top             =   6600
      Width           =   1125
   End
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   3945
      Left            =   2145
      TabIndex        =   41
      Top             =   1560
      Visible         =   0   'False
      Width           =   10350
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   3900
         Left            =   30
         TabIndex        =   42
         Top             =   30
         Width           =   10290
         _ExtentX        =   18150
         _ExtentY        =   6879
         _Version        =   393216
         AllowUpdate     =   -1  'True
         AllowArrows     =   -1  'True
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         HeadLines       =   1
         RowHeight       =   19
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
      Height          =   8010
      Left            =   15
      TabIndex        =   18
      Top             =   -60
      Width           =   14310
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Height          =   690
         Left            =   75
         TabIndex        =   19
         Top             =   0
         Width           =   14190
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
            Left            =   915
            TabIndex        =   1
            Top             =   225
            Width           =   885
         End
         Begin VB.PictureBox Picture3 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9645
            ScaleHeight     =   240
            ScaleWidth      =   2070
            TabIndex        =   62
            Top             =   195
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9555
            ScaleHeight     =   240
            ScaleWidth      =   2070
            TabIndex        =   61
            Top             =   390
            Visible         =   0   'False
            Width           =   2070
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            FillStyle       =   0  'Solid
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   9555
            ScaleHeight     =   240
            ScaleWidth      =   1095
            TabIndex        =   60
            Top             =   15
            Visible         =   0   'False
            Width           =   1095
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
            TabIndex        =   2
            Top             =   225
            Width           =   2115
         End
         Begin VB.Label LBLCUSTOMER 
            Height          =   390
            Left            =   12105
            TabIndex        =   70
            Top             =   255
            Width           =   1815
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   255
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         Height          =   5145
         Left            =   75
         TabIndex        =   25
         Top             =   600
         Width           =   14190
         Begin MSFlexGridLib.MSFlexGrid grdsales 
            Height          =   4905
            Left            =   75
            TabIndex        =   3
            Top             =   165
            Width           =   14055
            _ExtentX        =   24791
            _ExtentY        =   8652
            _Version        =   393216
            Rows            =   1
            Cols            =   19
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            HighLight       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            GridLineWidth   =   2
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         Height          =   2325
         Left            =   75
         TabIndex        =   26
         Top             =   5655
         Width           =   14190
         Begin VB.TextBox TxtBarcode 
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
            Height          =   375
            Left            =   630
            TabIndex        =   0
            Top             =   510
            Width           =   2025
         End
         Begin VB.TextBox Txtsize 
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
            Height          =   375
            Left            =   11895
            MaxLength       =   7
            TabIndex        =   56
            Top             =   1935
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.ComboBox cmbcolor 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            Left            =   13710
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   1365
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.ComboBox CmbPack 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   390
            ItemData        =   "FrmCrdNote.frx":0000
            Left            =   8745
            List            =   "FrmCrdNote.frx":001F
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   495
            Width           =   990
         End
         Begin VB.TextBox txttax 
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
            Height          =   390
            Left            =   11790
            MaxLength       =   6
            TabIndex        =   52
            Top             =   495
            Width           =   900
         End
         Begin VB.TextBox TxtSale_Rate 
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
            Height          =   390
            Left            =   10680
            MaxLength       =   6
            TabIndex        =   49
            Top             =   495
            Width           =   1095
         End
         Begin VB.TextBox TXTMFGR 
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
            TabIndex        =   48
            Top             =   1860
            Visible         =   0   'False
            Width           =   1920
         End
         Begin VB.TextBox Txtpack 
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
            Height          =   390
            Left            =   2925
            TabIndex        =   46
            Top             =   1755
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.TextBox txtratehide 
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
            Left            =   4530
            MaxLength       =   6
            TabIndex        =   45
            Top             =   1635
            Visible         =   0   'False
            Width           =   630
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
            Height          =   480
            Left            =   3795
            TabIndex        =   13
            Top             =   960
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
            Height          =   390
            Left            =   45
            TabIndex        =   4
            Top             =   490
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
            Height          =   390
            Left            =   4215
            TabIndex        =   5
            Top             =   490
            Width           =   3615
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
            Height          =   390
            Left            =   7845
            MaxLength       =   7
            TabIndex        =   7
            Top             =   490
            Width           =   900
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
            Height          =   390
            Left            =   9735
            MaxLength       =   6
            TabIndex        =   8
            Top             =   495
            Width           =   930
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
            Left            =   6240
            TabIndex        =   15
            Top             =   975
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
            Left            =   5055
            TabIndex        =   14
            Top             =   960
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
            Height          =   390
            Left            =   2670
            TabIndex        =   29
            Top             =   495
            Width           =   1530
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
            Height          =   390
            Left            =   10710
            MaxLength       =   15
            TabIndex        =   11
            Top             =   1920
            Visible         =   0   'False
            Width           =   1155
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            Height          =   390
            Left            =   3780
            TabIndex        =   6
            Top             =   1755
            Visible         =   0   'False
            Width           =   810
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
            Left            =   7440
            TabIndex        =   16
            Top             =   975
            Width           =   1125
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   390
            Left            =   5310
            TabIndex        =   9
            Top             =   1770
            Visible         =   0   'False
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   688
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
            Height          =   390
            Left            =   5070
            MaxLength       =   10
            TabIndex        =   10
            Top             =   1665
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label lbllineno 
            Height          =   240
            Left            =   12855
            TabIndex        =   76
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label lbltrxyear 
            Height          =   285
            Left            =   12855
            TabIndex        =   75
            Top             =   960
            Width           =   735
         End
         Begin VB.Label lbltrxtype 
            Height          =   315
            Left            =   10920
            TabIndex        =   74
            Top             =   990
            Width           =   675
         End
         Begin VB.Label lblvchno 
            Height          =   255
            Left            =   10125
            TabIndex        =   73
            Top             =   1050
            Width           =   615
         End
         Begin VB.Label lblcost 
            Height          =   390
            Left            =   11670
            TabIndex        =   72
            Top             =   915
            Width           =   1035
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Barcode"
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
            Height          =   300
            Index           =   22
            Left            =   630
            TabIndex        =   71
            Top             =   225
            Width           =   2025
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Color"
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
            Height          =   300
            Index           =   21
            Left            =   13710
            TabIndex        =   59
            Top             =   1080
            Visible         =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Size"
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
            Height          =   300
            Index           =   19
            Left            =   11895
            TabIndex        =   58
            Top             =   1650
            Visible         =   0   'False
            Width           =   810
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pack"
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
            Height          =   300
            Index           =   13
            Left            =   8760
            TabIndex        =   57
            Top             =   225
            Width           =   945
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Tax %"
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
            Height          =   300
            Index           =   12
            Left            =   11790
            TabIndex        =   53
            Top             =   225
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Item Code"
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
            Height          =   300
            Index           =   5
            Left            =   2670
            TabIndex        =   51
            Top             =   225
            Width           =   1530
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Sold Rate"
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
            Height          =   300
            Index           =   2
            Left            =   10680
            TabIndex        =   50
            Top             =   225
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Pack"
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
            Height          =   300
            Index           =   4
            Left            =   2925
            TabIndex        =   47
            Top             =   1485
            Visible         =   0   'False
            Width           =   720
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "NET AMOUNT"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   375
            Index           =   6
            Left            =   6465
            TabIndex        =   44
            Top             =   1740
            Width           =   1515
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   570
            Left            =   8070
            TabIndex        =   43
            Top             =   1605
            Width           =   1770
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
            Height          =   300
            Index           =   8
            Left            =   45
            TabIndex        =   40
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
            Height          =   300
            Index           =   9
            Left            =   4215
            TabIndex        =   39
            Top             =   225
            Width           =   3615
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
            Height          =   300
            Index           =   10
            Left            =   7845
            TabIndex        =   38
            Top             =   225
            Width           =   900
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "MRP"
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
            Height          =   300
            Index           =   11
            Left            =   9735
            TabIndex        =   37
            Top             =   225
            Width           =   930
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
            Height          =   300
            Index           =   14
            Left            =   12705
            TabIndex        =   36
            Top             =   225
            Width           =   1455
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
            TabIndex        =   35
            Top             =   1005
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
            Height          =   300
            Index           =   16
            Left            =   5310
            TabIndex        =   34
            Top             =   1500
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Article No."
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
            Height          =   300
            Index           =   7
            Left            =   10710
            TabIndex        =   33
            Top             =   1650
            Visible         =   0   'False
            Width           =   1155
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
            Height          =   390
            Left            =   12705
            TabIndex        =   12
            Top             =   495
            Width           =   1455
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
            TabIndex        =   32
            Top             =   1215
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Sell Unit"
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
            Height          =   300
            Index           =   20
            Left            =   3780
            TabIndex        =   31
            Top             =   1485
            Visible         =   0   'False
            Width           =   810
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
            TabIndex        =   30
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
Option Explicit
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT As Boolean
Dim PHY_PRERATE As New ADODB.Recordset
Dim PRERATE_FLAG As Boolean

'Dim CN_FLAG As Boolean

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            CmbPack.Enabled = False
            TXTRATE.Enabled = True
            TXTRATE.SetFocus
         Case vbKeyEscape
            'TXTUNIT.Text = ""
            CmbPack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
End Sub

Private Sub CMDADD_Click()
    Dim RSTTRXFILE, TRXMAST As ADODB.Recordset
    Dim i, LINNO As Integer
    
    LINNO = 1
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(LINE_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= '" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(txtBillNo.Text) & "", db, adOpenForwardOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        LINNO = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing

    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = Trim(TXTITEMCODE.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(TXTPRODUCT.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Val(TXTQTY.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Val(TXTTAX.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 5) = Format(Val(TXTRATE.Text), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 7) = Trim(txtBatch.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 8) = Val(TxtSale_Rate.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 9) = Format(Val(LBLSUBTOTAL.Caption), ".000")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 10) = txtBillNo.Text
    If M_EDIT = True Then
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
    Else
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = LINNO
    End If
    grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = Val(Txtpack.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = Trim(TXTMFGR.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = CmbPack.Text
    grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = Val(Txtsize.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = Trim(cmbcolor.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = Trim(TXTITEMCODE.Text) & Trim(Txtsize.Text) & cmbcolor.Text
    grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = Trim(TxtBarcode.Text)
    
    Dim R_RATE, W_RATE, V_RATE, CRTN, CRTN_PACK As Double
    
    Set RSTRTRXFILE = New ADODB.Recordset
    RSTRTRXFILE.Open "SELECT * From TRXSUB WHERE TRX_YEAR= '" & lbltrxyear.Caption & "' AND TRX_TYPE= '" & lbltrxtype.Caption & "' AND VCH_NO = " & Val(lblvchno.Caption) & " AND LINE_NO= " & Val(lbllineno.Caption) & "", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
        Dim RSTRTRXFILE2 As ADODB.Recordset
        Set RSTRTRXFILE2 = New ADODB.Recordset
        RSTRTRXFILE2.Open "SELECT * From RTRXFILE WHERE ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "' AND TRX_YEAR= '" & RSTRTRXFILE!R_TRX_YEAR & "' AND TRX_TYPE= '" & RSTRTRXFILE!R_TRX_TYPE & "' AND VCH_NO = " & RSTRTRXFILE!R_VCH_NO & " AND LINE_NO= " & RSTRTRXFILE!R_LINE_NO & "", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTRTRXFILE2.EOF And RSTRTRXFILE2.BOF) Then
            grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = IIf(IsNull(RSTRTRXFILE2!ITEM_COST), 0, RSTRTRXFILE2!ITEM_COST)
            R_RATE = IIf(IsNull(RSTRTRXFILE2!P_RETAIL), 0, RSTRTRXFILE2!P_RETAIL)
            W_RATE = IIf(IsNull(RSTRTRXFILE2!P_WS), 0, RSTRTRXFILE2!P_WS)
            V_RATE = IIf(IsNull(RSTRTRXFILE2!P_VAN), 0, RSTRTRXFILE2!P_VAN)
            CRTN = IIf(IsNull(RSTRTRXFILE2!P_CRTN), 0, RSTRTRXFILE2!P_CRTN)
            CRTN_PACK = IIf(IsNull(RSTRTRXFILE2!CRTN_PACK), 0, RSTRTRXFILE2!CRTN_PACK)
        End If
        RSTRTRXFILE2.Close
        Set RSTRTRXFILE2 = Nothing
    End If
    RSTRTRXFILE.Close
    Set RSTRTRXFILE = Nothing
    
    If Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) = 0 Then grdsales.TextMatrix(Val(TXTSLNO.Text), 6) = Format(Val(lblcost.Caption), ".000")
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            If R_RATE = 0 Then R_RATE = IIf(IsNull(!P_RETAIL), 0, !P_RETAIL)
            If W_RATE = 0 Then W_RATE = IIf(IsNull(!P_WS), 0, !P_WS)
            If V_RATE = 0 Then V_RATE = IIf(IsNull(!P_VAN), 0, !P_VAN)
            If CRTN = 0 Then CRTN = IIf(IsNull(!P_CRTN), 0, !P_CRTN)
            If CRTN_PACK = 0 Then CRTN_PACK = IIf(IsNull(!CRTN_PACK), 0, !CRTN_PACK)
            RSTTRXFILE.Update
        End If
    End With
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
'    CN_FLAG = False

'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.BAL_QTY > 0", db, adOpenStatic, adLockOptimistic, adCmdText
'    With RSTTRXFILE
'        If Not (.EOF And .BOF) Then
'            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
'            !ISSUE_QTY = !ISSUE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
'
'            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
'            !BAL_QTY = !BAL_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
'
'            RSTTRXFILE.Update
'        Else
''            CN_FLAG = True
'        End If
'    End With
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    db.Execute "delete From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= '" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND ITEM_CODE='" & Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1)) & "'AND LINE_NO=" & Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11)) & ""
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With RSTTRXFILE
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = Trim(creditbill.LBLTYPE.Caption)
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.value)
        RSTTRXFILE!VCH_NO = Trim(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(Trim(LBLDATE.Caption), "dd/mm/yyyy")
        RSTTRXFILE!LINE_NO = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
        RSTTRXFILE!Category = "GENERAL"
        RSTTRXFILE!ITEM_CODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 1))
        RSTTRXFILE!ITEM_NAME = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 2))
        RSTTRXFILE!QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
        RSTTRXFILE!UNIT = 1
        RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
        RSTTRXFILE!MRP = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 5))
        RSTTRXFILE!ITEM_COST = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
        RSTTRXFILE!PTR = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6))
        RSTTRXFILE!SALES_PRICE = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
        'RSTTRXFILE!SALES_TAX = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
        RSTTRXFILE!VCH_DESC = "Sales Return"
        RSTTRXFILE!REF_NO = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 7))
        'RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) '* Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
'        If CN_FLAG = True Then
'            RSTTRXFILE!BAL_QTY = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3)) '* Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
'        Else
'            RSTTRXFILE!BAL_QTY = 0
'        End If
        RSTTRXFILE!P_RETAIL = R_RATE
        RSTTRXFILE!P_WS = W_RATE
        RSTTRXFILE!P_CRTN = CRTN
        RSTTRXFILE!P_VAN = V_RATE
        RSTTRXFILE!CRTN_PACK = CRTN_PACK
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 9))
        RSTTRXFILE!LINE_DISC = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 12))
        RSTTRXFILE!MFGR = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
        RSTTRXFILE!BARCODE = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
        RSTTRXFILE!LOOSE_PACK = 1
        RSTTRXFILE!PACK_TYPE = CmbPack.Text
        
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Date
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "311000"
        RSTTRXFILE!CHECK_FLAG = "V"
        RSTTRXFILE!PINV = "CN# " & Trim(txtBillNo.Text)
        
        RSTTRXFILE.Update
    End With
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 9)), ".00")
    Next i
    
    
    Picture1.Cls
    Picture1.Picture = Nothing
    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
    Picture1.Print TXTPRODUCT.Text
    
    'Picture3.ScaleMode = 1                               'pixels
    
    cmdadd.Tag = Val(LBLTOTAL.Caption)
    TXTSLNO.Text = grdsales.Rows
    TXTPRODUCT.Text = ""
    
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTUNIT.Text = ""
    Txtpack.Text = ""
    TXTTAX.Text = ""
    TxtBarcode.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TxtSale_Rate.Text = ""
    txtratehide.Text = ""
    txtBatch.Text = ""
    TXTMFGR.Text = ""
    TXTEXPDATE.Text = ""
    TXTEXPIRY.Text = "  /  "
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
    M_EDIT = False
    TXTSLNO.Enabled = False
    TxtBarcode.Enabled = True
    TxtBarcode.SetFocus
    txtBillNo.Enabled = False
    cmdRefresh.Enabled = True
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
            Exit Sub
    End Select
End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    db.Execute "DELETE From RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "' AND TRX_TYPE= '" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' "
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
        grdsales.TextMatrix(Val(TXTSLNO.Text), 11) = grdsales.TextMatrix(i + 1, 11)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 12) = grdsales.TextMatrix(i + 1, 12)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 13) = grdsales.TextMatrix(i + 1, 13)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 14) = grdsales.TextMatrix(i + 1, 14)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 15) = grdsales.TextMatrix(i + 1, 15)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 16) = grdsales.TextMatrix(i + 1, 16)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 17) = grdsales.TextMatrix(i + 1, 17)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 18) = grdsales.TextMatrix(i + 1, 18)
    Next i
    grdsales.Rows = grdsales.Rows - 1
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.Rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 9)), ".00")
    Next i
    
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTUNIT.Text = ""
    Txtpack.Text = ""
    TXTQTY.Text = ""
    TXTRATE.Text = ""
    TxtSale_Rate.Text = ""
    txtratehide.Text = ""
    TXTEXPDATE.Text = ""
    txtBatch.Text = ""
    TXTMFGR.Text = ""
    LBLSUBTOTAL.Caption = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CmdExit.Enabled = False
    
    If grdsales.Rows = 1 Then
        CmdExit.Enabled = True
    End If
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()

    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    Dim RSTTRXFILE As ADODB.Recordset
    
    If Val(TXTSLNO.Text) >= grdsales.Rows Then Exit Sub
    
    On Error GoTo eRRhAND
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!ISSUE_VAL)) Then !ISSUE_VAL = 0
            !ISSUE_VAL = !ISSUE_VAL + Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            !CLOSE_QTY = !CLOSE_QTY - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
            If (IsNull(!CLOSE_VAL)) Then !CLOSE_VAL = 0
            !CLOSE_VAL = !CLOSE_VAL - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 11))
            RSTTRXFILE.Update
        End If
    End With
    db.CommitTrans
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
       
    db.Execute "DELETE From RTRXFILE WHERE ITEM_CODE = '" & grdsales.TextMatrix(Val(TXTSLNO.Text), 1) & "' AND TRX_TYPE= '" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' "
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CmdExit.Enabled = False
    TXTQTY.Enabled = True
    TXTQTY.SetFocus
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TxtBarcode.Text = ""
            TXTPRODUCT.Text = ""
            TXTQTY.Text = ""
            TXTRATE.Text = ""
            TxtSale_Rate.Text = ""
            txtratehide.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTITEMCODE.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
            TXTUNIT.Text = ""
            Txtpack.Text = ""
            LBLSUBTOTAL.Caption = ""
            TXTEXPDATE.Text = ""
            txtBatch.Text = ""
            TXTMFGR.Text = ""
            Frame4.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TxtSale_Rate.Enabled = False
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            txtBatch.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    On Error GoTo eRRhAND
    CLOSEALL = 0
    Unload Me
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    LBLDATE.Caption = Date
    lbltime.Caption = Time
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 0
    grdsales.ColWidth(2) = 2700
    grdsales.ColWidth(3) = 800
    grdsales.ColWidth(4) = 900
    grdsales.ColWidth(5) = 1000
    grdsales.ColWidth(6) = 1000
    grdsales.ColWidth(7) = 1000
    grdsales.ColWidth(8) = 0
    grdsales.ColWidth(9) = 1300
    grdsales.ColWidth(10) = 0
    grdsales.ColWidth(11) = 0
    grdsales.ColWidth(12) = 1100
    grdsales.ColWidth(13) = 1100
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ITEM CODE"
    grdsales.TextArray(2) = "ITEM NAME"
    grdsales.TextArray(3) = "QTY"
    grdsales.TextArray(4) = "Tax"
    grdsales.TextArray(5) = "MRP"
    grdsales.TextArray(6) = "RATE"
    grdsales.TextArray(7) = "BATCH"
    grdsales.TextArray(8) = "EXPIRY"
    grdsales.TextArray(9) = "SUB TOTAL"
    grdsales.TextArray(10) = "Vch No"
    grdsales.TextArray(11) = "Line No"
    grdsales.TextArray(12) = "Pack"
    grdsales.TextArray(13) = "MFGR"

    PHYFLAG = True
    PRERATE_FLAG = True
    TXTPRODUCT.Enabled = False
    TxtBarcode.Enabled = False
    TXTQTY.Enabled = False
    TXTRATE.Enabled = False
    TxtSale_Rate.Enabled = False
    TXTEXPDATE.Enabled = False
    txtBatch.Enabled = False
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTUNIT.Enabled = False
    TXTREMARKS.Text = "SALES RETURN"
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    CLOSEALL = 1
    
    LBLTOTAL.Caption = 0
    
    Dim TRXMAST As ADODB.Recordset
    If creditbill.TxtCN.Text <> 0 Then
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select VCH_NO From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE= '" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(creditbill.TxtCN.Text) & " ", db, adOpenForwardOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0))
            TRXMAST.Close
            Set TRXMAST = Nothing
        Else
            Set TRXMAST = New ADODB.Recordset
            TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = '" & Trim(creditbill.LBLTYPE.Caption) & "'", db, adOpenForwardOnly
            If Not (TRXMAST.EOF And TRXMAST.BOF) Then
                txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
            End If
            TRXMAST.Close
            Set TRXMAST = Nothing
        End If
    Else
        Set TRXMAST = New ADODB.Recordset
        TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE = '" & Trim(creditbill.LBLTYPE.Caption) & "'", db, adOpenForwardOnly
        If Not (TRXMAST.EOF And TRXMAST.BOF) Then
            txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        End If
        TRXMAST.Close
        Set TRXMAST = Nothing
    End If
    Call TXTBILLNO_KeyDown(13, 0)
        
'    Width = 9660
'    Height = 8300
    cetre Me
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If PRERATE_FLAG = False Then PHY_PRERATE.Close
        If Tag = "Y" Then
            creditbill.Enabled = True
            MDIMAIN.Enabled = True
            creditbill.LBLRETAMT.Caption = Format(Round(Val(LBLTOTAL.Caption), 2), "0.00")
            creditbill.lblnetamount.Caption = Round(Val(creditbill.LBLTOTAL.Caption) - (Val(creditbill.TXTAMOUNT.Text) + Val(creditbill.LBLRETAMT.Caption)), 2) + Val(creditbill.LBLFOT.Caption) + Val(creditbill.TxtFrieght.Text)
            'creditbill.lblnetamount.Caption = Format(Val(creditbill.LBLTOTAL.Caption) - (Val(creditbill.LBLDISCAMT.Caption) + Val(creditbill.LBLRETAMT.Caption)), ".00")
'            If creditbill.TXTSLNO.Enabled = True Then creditbill.TXTSLNO.SetFocus
'            If creditbill.TXTPRODUCT.Enabled = True Then creditbill.TXTPRODUCT.SetFocus
'            If creditbill.TXTQTY.Enabled = True Then creditbill.TXTQTY.SetFocus
'            If creditbill.TXTRATE.Enabled = True Then creditbill.TXTRATE.SetFocus
'            If creditbill.TXTTAX.Enabled = True Then creditbill.TXTTAX.SetFocus
'            If creditbill.TXTEXPIRY.Visible = True Then creditbill.TXTEXPIRY.SetFocus
'            If creditbill.txtexpdate.Enabled = True Then creditbill.txtexpdate.SetFocus
'            If creditbill.txtBatch.Enabled = True Then creditbill.txtBatch.SetFocus
'            If creditbill.TXTDISC.Enabled = True Then creditbill.TXTDISC.SetFocus
'            If creditbill.cmdadd.Enabled = True Then creditbill.cmdadd.SetFocus
            If grdsales.Rows > 1 Then
                creditbill.TxtCN.Text = Val(txtBillNo.Text)
                creditbill.CmdExit.Enabled = False
            Else
                creditbill.TxtCN.Text = ""
            End If
            
        Else
            MDIMAIN.Enabled = True
            MDIMAIN.PCTMENU.SetFocus
        End If
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDPRERATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTQTY.Text = GRDPRERATE.Columns(4)
            TXTRATE.Text = GRDPRERATE.Columns(14)
            TxtSale_Rate.Text = GRDPRERATE.Columns(5)
            TXTTAX.Text = GRDPRERATE.Columns(7)
            lblcost.Caption = IIf(IsNull(GRDPRERATE.Columns(10)), "", GRDPRERATE.Columns(10))
            'TxtSale_Rate.Text = Round(Val(TxtSale_Rate.Text) * 100 / (Val(TXTTAX.Text) + 100), 3)
            txtBatch.Text = GRDPRERATE.Columns(8)
            'Txtsize.Text = GRDPRERATE.Columns(9)
            
            lblvchno.Caption = GRDPRERATE.Columns(11)
            lbltrxtype.Caption = GRDPRERATE.Columns(0)
            lbltrxyear.Caption = GRDPRERATE.Columns(13)
            lbllineno.Caption = GRDPRERATE.Columns(9)
            
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
    End Select
'    Select Case KeyCode
'        Case vbKeyEscape
'            Set GRDPRERATE.DataSource = Nothing
'            fRMEPRERATE.Visible = False
'            Fram.Enabled = True
'            cmbpac.Enabled = True
'            TXTPTR.SetFocus
'    End Select
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn

            TXTPRODUCT.Text = grdtmp.Columns(1)
            TXTITEMCODE.Text = grdtmp.Columns(0)
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Call FILL_PREVIIOUSRATE
            Exit Sub
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'  ORDER BY TRX_TYPE desc, VCH_NO DESC", db, adOpenForwardOnly
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                TXTUNIT.Text = RSTRXFILE!UNIT
                Txtpack.Text = RSTRXFILE!LINE_DISC
                txtratehide.Text = RSTRXFILE!MRP
                TxtSale_Rate.Text = RSTRXFILE!SALES_PRICE
                TXTRATE.Text = RSTRXFILE!MRP
                'TXTEXPIRY.Text = Format(RSTRXFILE!EXP_DATE, "MM/YY")
                'txtexpdate.Text = RSTRXFILE!EXP_DATE
                txtBatch.Text = RSTRXFILE!REF_NO
                TXTMFGR.Text = IIf(IsNull(RSTRXFILE!MFGR), "", RSTRXFILE!MFGR)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), CmbPack.ListIndex = -1, RSTRXFILE!PACK_TYPE)
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            Set grdtmp.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            Fram.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TXTUNIT.Text = ""
            Txtpack.Text = ""
            TXTVCHNO.Text = ""
            TXTLINENO.Text = ""
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
            'If Trim(txtBatch.Text) = "" Then Exit Sub
            txtBatch.Enabled = False
            Txtsize.Enabled = True
            Txtsize.SetFocus
        Case vbKeyEscape
            'txtexpdate.Enabled = True
            TxtSale_Rate.Enabled = True
            txtBatch.Enabled = False
            TxtSale_Rate.SetFocus
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
Dim rstTRXMAST As ADODB.Recordset
Dim i As Long

    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txtBillNo.Text) = 0 Then Exit Sub
            grdsales.Rows = 1
           
            i = 0
            grdsales.Rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From RTRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND TRX_TYPE='" & Trim(creditbill.LBLTYPE.Caption) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY LINE_NO", db, adOpenForwardOnly
            Do Until rstTRXMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ITEM_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ITEM_NAME
                grdsales.TextMatrix(i, 3) = rstTRXMAST!QTY
                grdsales.TextMatrix(i, 4) = IIf(IsNull(rstTRXMAST!SALES_TAX), "", rstTRXMAST!SALES_TAX)
                grdsales.TextMatrix(i, 5) = Format(rstTRXMAST!MRP, ".000")
                grdsales.TextMatrix(i, 6) = Format(rstTRXMAST!ITEM_COST, ".000")
                grdsales.TextMatrix(i, 7) = rstTRXMAST!REF_NO
                grdsales.TextMatrix(i, 8) = Format(rstTRXMAST!SALES_PRICE, ".000") 'Format(rstTRXMAST!EXP_DATE, "DD/MM/YYYY")
                grdsales.TextMatrix(i, 9) = Format(rstTRXMAST!TRX_TOTAL, ".000")
                
                grdsales.TextMatrix(i, 10) = rstTRXMAST!VCH_NO
                grdsales.TextMatrix(i, 11) = rstTRXMAST!LINE_NO
                grdsales.TextMatrix(i, 12) = rstTRXMAST!LINE_DISC
                TXTREMARKS.Text = Mid(rstTRXMAST!VCH_DESC, 12)
                grdsales.TextMatrix(i, 13) = IIf(IsNull(rstTRXMAST!MFGR), "", rstTRXMAST!MFGR)
                grdsales.TextMatrix(i, 14) = IIf(IsNull(rstTRXMAST!PACK_TYPE), "Nos", rstTRXMAST!PACK_TYPE)
                grdsales.TextMatrix(i, 18) = IIf(IsNull(rstTRXMAST!BARCODE), "", rstTRXMAST!BARCODE)
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            LBLTOTAL.Caption = ""
            For i = 1 To grdsales.Rows - 1
                grdsales.TextMatrix(i, 0) = i
                LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 9)), ".00")
            Next i
            
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = False
            TxtBarcode.Enabled = True
            On Error Resume Next
            TxtBarcode.SetFocus
            On Error GoTo eRRhAND
            txtBillNo.Enabled = False
            Frame1.Enabled = False
            Frame4.Enabled = True
            
            'TXTSLNO.SetFocus
            If grdsales.Rows > 1 Then
                cmdRefresh.Enabled = True
                CmdExit.Enabled = False
                'Frame1.Enabled = False
                'Frame4.Enabled = False
                'CMDEXIT.Caption = "&CANCEL"
                'CMDEXIT.SetFocus
            Else
                'TXTREMARKS.SetFocus
            End If

            
    End Select
    
    Exit Sub
eRRhAND:
    MsgBox Err.Description
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
Dim TRXMAST As ADODB.Recordset
Dim i As Double

i = 1
Set TRXMAST = New ADODB.Recordset
TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = '" & Trim(creditbill.LBLTYPE.Caption) & "'", db, adOpenForwardOnly
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
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not IsDate(TXTEXPDATE.Text) Then Exit Sub
            txtBatch.Enabled = True
            TXTEXPIRY.Visible = False
            TXTEXPDATE.Enabled = False
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTRATE.Enabled = True
            TXTEXPDATE.Enabled = False
            TXTEXPIRY.Visible = False
            TXTRATE.SetFocus
    End Select
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTRXFILE As ADODB.Recordset
Dim i As Long
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
             If Trim(TXTPRODUCT.Text) = "" Then
                TXTITEMCODE.Enabled = True
                TXTPRODUCT.Enabled = False
                TXTITEMCODE.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select ITEM_CODE,ITEM_NAME, CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select ITEM_CODE,ITEM_NAME, CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
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
                TXTPRODUCT.Text = grdtmp.Columns(1)
                TXTITEMCODE.Text = grdtmp.Columns(0)
                Set grdtmp.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_PREVIIOUSRATE
                Exit Sub
            
'                Set RSTRXFILE = New ADODB.Recordset
'                RSTRXFILE.Open "Select * From TRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'  ORDER BY TRX_TYPE desc, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
'                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
'                    TXTUNIT.Text = RSTRXFILE!UNIT
'                    Txtpack.Text = RSTRXFILE!UNIT
'                    txtratehide.Text = RSTRXFILE!MRP
'                    TxtSale_Rate.Text = IIf(IsNull(RSTRXFILE!P_RETAILWOTAX), "", RSTRXFILE!P_RETAILWOTAX)
'                    TXTRATE.Text = IIf(IsNull(RSTRXFILE!MRP), "", RSTRXFILE!MRP)
'                    TXTTAX.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", RSTRXFILE!SALES_TAX)
'                    'TXTEXPIRY.Text = Format(RSTRXFILE!EXP_DATE, "MM/YY")
'                    txtBatch.Text = IIf(IsNull(RSTRXFILE!REF_NO), "", RSTRXFILE!REF_NO)
'                    TXTMFGR.Text = IIf(IsNull(RSTRXFILE!MFGR), "", RSTRXFILE!MFGR)
'                    On Error Resume Next
'                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), CmbPack.ListIndex = -1, RSTRXFILE!PACK_TYPE)
'                    On Error GoTo eRRhAND
'                End If
'                RSTRXFILE.Close
'                Set RSTRXFILE = Nothing
            
'                If PHY.RecordCount = 1 Then
'                    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
'                    TXTPRODUCT.Enabled = False
'                    TXTQTY.Enabled = True
'                    TXTQTY.SetFocus
'                    Exit Sub
'                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 5000
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1100
                grdtmp.SetFocus
                'Exit Sub
            End If
        Case vbKeyEscape
            TXTITEMCODE.Enabled = True
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTITEMCODE.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
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
            CmbPack.Enabled = True
            CmbPack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            TXTQTY.Enabled = False
            TxtBarcode.Enabled = True
            TxtBarcode.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_LostFocus()
    TXTQTY.Text = Format(TXTQTY.Text, ".00")
    Call TxtSale_Rate_LostFocus
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TxtSale_Rate.Text), 2)), ".000")
End Sub

Private Sub TXTRATE_GotFocus()
    'TXTRATE.Text = txtratehide.Text
    TXTRATE.SelStart = 0
    TXTRATE.SelLength = Len(TXTRATE.Text)
End Sub

Private Sub TXTRATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(TXTRATE.Text) = 0 Then Exit Sub
            TXTRATE.Enabled = False
            TxtSale_Rate.Enabled = True
            TxtSale_Rate.SetFocus
        Case vbKeyEscape
            CmbPack.Enabled = True
            TXTRATE.Enabled = False
            CmbPack.SetFocus
    End Select
End Sub

Private Sub TXTRATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTRATE_LostFocus()
    txtratehide.Text = Val(TXTRATE.Text)
    'TXTRATE.Text = Format(Val(TXTRATE.Text) / Val(TXTUNIT.Text), ".000")
    'TXTRATE.Text = Format(TXTRATE.Text, ".000")
    'LBLSUBTOTAL.Caption = Format((Val(TXTQTY.Text) * Round(Val(TXTRATE.Text), 2)), ".000")
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Set GRDPRERATE.DataSource = Nothing
    fRMEPRERATE.Visible = False
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
                TXTTAX.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 4)
                TXTUNIT.Text = "1"
                TXTRATE.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                TxtSale_Rate.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 6)
                txtratehide.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 5)
                txtBatch.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 7)
                lblcost.Caption = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 8))
                'txtexpdate.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 8)
                'TXTEXPIRY.Text = Format(grdsales.TextMatrix(Val(TXTSLNO.Text), 8), "mm/yy")
                LBLSUBTOTAL.Caption = grdsales.TextMatrix(Val(TXTSLNO.Text), 9)
                
                TXTVCHNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 10)
                TXTLINENO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 11)
                Txtpack.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 12)
                TXTMFGR.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 13))
                On Error Resume Next
                CmbPack.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 14))
                On Error GoTo eRRhAND
                TxtBarcode.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 18))
                
                TXTSLNO.Enabled = False
                TxtBarcode.Enabled = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = False
                TXTRATE.Enabled = False
                TxtSale_Rate.Enabled = False
                TXTEXPDATE.Enabled = False
                txtBatch.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TxtBarcode.Enabled = True
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TxtSale_Rate.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TxtBarcode.SetFocus
        Case vbKeyEscape
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTPRODUCT.Text = ""
                TXTITEMCODE.Text = ""
                TXTVCHNO.Text = ""
                TXTLINENO.Text = ""
                TXTUNIT.Text = ""
                Txtpack.Text = ""
                TXTQTY.Text = ""
                TXTRATE.Text = ""
                TxtSale_Rate.Text = ""
                txtratehide.Text = ""
                LBLSUBTOTAL.Caption = ""
                TXTEXPDATE.Text = ""
                txtBatch.Text = ""
                TXTMFGR.Text = ""
                cmbcolor.ListIndex = -1
                CmbPack.ListIndex = -1
                Txtsize.Text = ""
                TxtBarcode.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
                cmdRefresh.Enabled = True
                cmdRefresh.SetFocus
            End If
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description, vbOKOnly, "EzBiz"
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
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = True
            txtBatch.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            TxtSale_Rate.Enabled = True
            TXTEXPDATE.Enabled = False
            TxtSale_Rate.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPIRY_LostFocus()
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) = 0 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    If Val(Mid(TXTEXPIRY.Text, 4, 5)) = 0 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    
    If Val(Mid(TXTEXPIRY.Text, 1, 2)) > 12 Then
        TXTEXPDATE.Text = ""
        Exit Sub
    End If
    
    M = Val(Mid(TXTEXPIRY.Text, 1, 2))
    Y = Val(Right(TXTEXPIRY.Text, 2))
    Y = 2000 + Y
    M_DATE = "01" & "/" & M & "/" & Y
    D = LastDayOfMonth(M_DATE)
    M_DATE = D & "/" & M & "/" & Y
    TXTEXPDATE.Text = M_DATE
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
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

Private Sub TXTTAX_LostFocus()
    Call TxtSale_Rate_LostFocus
End Sub

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            
            TXTUNIT.Enabled = False
            Txtpack.Enabled = True
            Txtpack.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            TXTUNIT.Text = ""
            Txtpack.Text = ""
            TXTPRODUCT.Enabled = True
            TXTUNIT.Enabled = False
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtpack_GotFocus()
    Txtpack.SelStart = 0
    Txtpack.SelLength = Len(Txtpack.Text)
End Sub

Private Sub Txtpack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtpack.Text) = 0 Then Exit Sub
            Txtpack.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
         Case vbKeyEscape
            If M_EDIT = True Then Exit Sub
            'TXTUNIT.Text = ""
            Txtpack.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
    End Select
End Sub

Private Sub Txtpack_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtSale_Rate_GotFocus()
    TxtSale_Rate.SelStart = 0
    TxtSale_Rate.SelLength = Len(TxtSale_Rate.Text)
End Sub

Private Sub TxtSale_Rate_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TxtSale_Rate.Text) = 0 Then Exit Sub
            TxtSale_Rate.Enabled = False
            TXTTAX.Enabled = True
            TXTTAX.SetFocus
            'TXTEXPIRY.Visible = True
            'TXTEXPIRY.SetFocus
        Case vbKeyEscape
            TXTRATE.Enabled = True
            TxtSale_Rate.Enabled = False
            TXTRATE.SetFocus
    End Select
End Sub

Private Sub TxtSale_Rate_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtSale_Rate_LostFocus()
    TxtSale_Rate.Text = Format(Val(TxtSale_Rate.Text), ".000")
    'TxtSale_Rate.Text = Format(TxtSale_Rate.Text, ".000")
    LBLSUBTOTAL.Caption = Format(Round((Val(TxtSale_Rate.Text) * Val(TXTTAX.Text) / 100 + Val(TxtSale_Rate.Text)) * Val(TXTQTY.Text), 0), "0.000")
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTRXFILE As ADODB.Recordset
Dim i As Long
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTITEMCODE.Text) = "" Then
                TXTPRODUCT.Enabled = True
                TXTITEMCODE.Enabled = False
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
                
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ", db, adOpenForwardOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY From ITEMMAST  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' ", db, adOpenForwardOnly
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
                RSTRXFILE.Open "Select * From RTRXFILE  WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'  ORDER BY TRX_TYPE desc, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                    TXTUNIT.Text = RSTRXFILE!UNIT
                    Txtpack.Text = RSTRXFILE!LINE_DISC
                    txtratehide.Text = RSTRXFILE!MRP
                    TxtSale_Rate.Text = RSTRXFILE!SALES_PRICE
                    TXTRATE.Text = RSTRXFILE!MRP
                    'TXTEXPIRY.Text = Format(RSTRXFILE!EXP_DATE, "MM/YY")
                    txtBatch.Text = RSTRXFILE!REF_NO
                    TXTMFGR.Text = IIf(IsNull(RSTRXFILE!MFGR), "", RSTRXFILE!MFGR)
                End If
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
            
                If PHY.RecordCount = 1 Then
                    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
                    TXTITEMCODE.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    Exit Sub
                End If
            ElseIf PHY.RecordCount > 1 Then
                FRMEGRDTMP.Visible = True
                Fram.Enabled = False
                grdtmp.Columns(0).Visible = False
                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
                grdtmp.Columns(1).Width = 5000
                'grdtmp.Columns(2).Visible = False
                grdtmp.Columns(2).Caption = "QTY"
                grdtmp.Columns(2).Width = 1100
                grdtmp.SetFocus
            End If
            
        Case vbKeyEscape
            TxtBarcode.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TxtBarcode.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtTax_GotFocus()
    TXTTAX.SelStart = 0
    TXTTAX.SelLength = Len(TXTTAX.Text)
End Sub

Private Sub TxtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Val(txttax.Text) = 0 Then Exit Sub
            TXTTAX.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TxtSale_Rate.Enabled = True
            TXTTAX.Enabled = False
            TxtSale_Rate.SetFocus
    End Select
End Sub

Private Sub TxtTax_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtsize_GotFocus()
    Txtsize.SelStart = 0
    Txtsize.SelLength = Len(Txtsize.Text)
End Sub

Private Sub Txtsize_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txtsize.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            txtBatch.Enabled = True
            Txtsize.Enabled = False
            txtBatch.SetFocus
    End Select
End Sub

Private Sub Txtsize_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub cmbcolor_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmbcolor.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
         Case vbKeyEscape
            cmbcolor.Enabled = False
            Txtsize.Enabled = True
            Txtsize.SetFocus
    End Select
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.Text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
Dim RSTRXFILE As ADODB.Recordset
Dim i As Long
    On Error GoTo eRRhAND
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TxtBarcode.Text) = "" Then
                TXTPRODUCT.Enabled = True
                TxtBarcode.Enabled = False
                TXTPRODUCT.SetFocus
                Exit Sub
            End If
            CmdDelete.Enabled = False
            
            Set RSTRXFILE = New ADODB.Recordset
            RSTRXFILE.Open "Select * From TRXFILE  WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "'  ORDER BY VCH_DATE DESC, TRX_TYPE DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
            If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
                TXTITEMCODE.Text = RSTRXFILE!ITEM_CODE
                TXTPRODUCT.Text = RSTRXFILE!ITEM_NAME
                For i = 1 To grdsales.Rows - 1
                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
                    End If
                Next i
                TXTUNIT.Text = RSTRXFILE!UNIT
                Txtpack.Text = RSTRXFILE!LINE_DISC
                txtratehide.Text = RSTRXFILE!MRP
                TxtSale_Rate.Text = RSTRXFILE!P_RETAILWOTAX
                TXTRATE.Text = RSTRXFILE!MRP
                TXTTAX.Text = IIf(IsNull(RSTRXFILE!SALES_TAX), "", RSTRXFILE!SALES_TAX)
                'TXTEXPIRY.Text = Format(RSTRXFILE!EXP_DATE, "MM/YY")
                txtBatch.Text = RSTRXFILE!REF_NO
                TXTMFGR.Text = IIf(IsNull(RSTRXFILE!MFGR), "", RSTRXFILE!MFGR)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), CmbPack.ListIndex = -1, RSTRXFILE!PACK_TYPE)
                On Error GoTo eRRhAND
                Call FILL_PRERATEWITHBARCODE
            Else
                MsgBox "ITEM NOT FOUND!!!!", vbOKOnly, "SALES RETURN"
                RSTRXFILE.Close
                Set RSTRXFILE = Nothing
                'TxtBarcode.Enabled = False
                'TXTPRODUCT.Enabled = True
                'TXTPRODUCT.SetFocus
                Exit Sub
            End If
            RSTRXFILE.Close
            Set RSTRXFILE = Nothing
            'TxtBarcode.Enabled = False
            'TXTQTY.Enabled = True
            'TXTQTY.SetFocus
'            Set grdtmp.DataSource = Nothing
'            If PHYFLAG = True Then
'                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY, BARCODE From ITEMMAST  WHERE BARCODE = '" & TxtBarcode.Text & "' ", db, adOpenForwardOnly
'                PHYFLAG = False
'            Else
'                PHY.Close
'                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY, BARCODE From ITEMMAST  WHERE BARCODE = '" & TxtBarcode.Text & "' ", db, adOpenForwardOnly
'                PHYFLAG = False
'            End If
'
'            Set grdtmp.DataSource = PHY
'            If PHY.RecordCount = 1 Then
'                TXTITEMCODE.Text = grdtmp.Columns(0)
'                TXTPRODUCT.Text = grdtmp.Columns(1)
'                For i = 1 To grdsales.Rows - 1
'                    If Trim(grdsales.TextMatrix(i, 1)) = Trim(TXTITEMCODE.Text) Then
'                        If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo, "SALES RETURN..") = vbNo Then Exit Sub
'                    End If
'                Next i
'
'                Set RSTRXFILE = New ADODB.Recordset
'                RSTRXFILE.Open "Select * From RTRXFILE  WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "'  ORDER BY TRX_TYPE desc, VCH_NO DESC", db, adOpenStatic, adLockReadOnly
'                If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
'                    TXTUNIT.Text = RSTRXFILE!UNIT
'                    Txtpack.Text = RSTRXFILE!LINE_DISC
'                    txtratehide.Text = RSTRXFILE!MRP
'                    TxtSale_Rate.Text = RSTRXFILE!SALES_PRICE
'                    TXTRATE.Text = RSTRXFILE!MRP
'                    'TXTEXPIRY.Text = Format(RSTRXFILE!EXP_DATE, "MM/YY")
'                    txtBatch.Text = RSTRXFILE!REF_NO
'                    TXTMFGR.Text = IIf(IsNull(RSTRXFILE!MFGR), "", RSTRXFILE!MFGR)
'                    On Error Resume Next
'                    CmbPack.Text = IIf(IsNull(RSTRXFILE!PACK_TYPE), CmbPack.ListIndex = -1, RSTRXFILE!PACK_TYPE)
'                    cmbcolor.Text = IIf(IsNull(RSTRXFILE!ITEM_COLOR), cmbcolor.ListIndex = -1, RSTRXFILE!ITEM_COLOR)
'                    On Error GoTo errHand
'                    Txtsize.Text = IIf(IsNull(RSTRXFILE!ITEM_SIZE), "", RSTRXFILE!ITEM_SIZE)
'                    TxtBarcode.Text = IIf(IsNull(RSTRXFILE!BARCODE), Trim(TXTITEMCODE.Text), RSTRXFILE!BARCODE)
'                End If
'                RSTRXFILE.Close
'                Set RSTRXFILE = Nothing
'
'                If PHY.RecordCount = 1 Then
'                    If Val(Txtpack.Text) = 0 Then Txtpack.Text = 1
'                    TXTITEMCODE.Enabled = False
'                    TXTQTY.Enabled = True
'                    TXTQTY.SetFocus
'                    Exit Sub
'                End If
'            ElseIf PHY.RecordCount > 1 Then
'                FRMEGRDTMP.Visible = True
'                Fram.Enabled = False
'                grdtmp.Columns(0).Visible = False
'                grdtmp.Columns(1).Caption = "PRODUCT DESCRIPTION"
'                grdtmp.Columns(1).Width = 3000
'                'grdtmp.Columns(2).Visible = False
'                grdtmp.Columns(2).Caption = "QTY"
'                grdtmp.Columns(2).Width = 1100
'                grdtmp.SetFocus
'            End If
            
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTITEMCODE.Enabled = False
            TXTQTY.Enabled = False
            TXTRATE.Enabled = False
            TXTEXPDATE.Enabled = False
            txtBatch.Enabled = False
            TXTSLNO.SetFocus
            CmdDelete.Enabled = False
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TxtBarcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Function print_labels()
    Dim wid As Single
    Dim hgt As Single
    Dim i As Long
    
    On Error GoTo eRRhAND
    
'    Dim P, PNAME
'    Dim printerfound As Boolean
'    printerfound = False
'    For Each P In Printers
'        PNAME = P.DeviceName
'        If UCase(Right(PNAME, 16)) Like "BAR CODE PRINTER" Then
'            Set Printer = P
'            printerfound = True
'            Exit For
'        End If
'    Next P
'    If printerfound = False Then
'        MsgBox ("Printer not found. Please correct the printer name")
'        Exit Function
'    End If
    
    
    i = Val(TXTQTY.Text)
     Do Until i <= 0
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture3.ScaleMode = vbPixels
        Picture4.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
        Printer.PaintPicture Picture3.Image, 300, 100 ', wid, hgt
        Printer.PaintPicture Picture4.Image, 1900, 100 ', wid, hgt
        Printer.PaintPicture Picture2.Image, 300, 310 ', wid, hgt
        Printer.PaintPicture Picture6.Image, 1000, 890 ', wid, hgt
        Printer.PaintPicture Picture1.Image, 300, 1100 ', wid, hgt
        Printer.PaintPicture Picture5.Image, 300, 1270 ', wid, hgt
       'wid = ScaleX(Picture2.ScaleWidth, Picture2.ScaleMode, _
        Printer.ScaleMode)
        'hgt = ScaleY(Picture2.ScaleHeight, _
        Picture2.ScaleMode, Printer.ScaleMode)
    '
    '    ' Draw the box.
        'Printer.Line (500, 500)-Step(wid, hgt), , B
        
        ' Finish printing.
        'Printer.EndDoc
        
        'Picture2.ScaleMode = vbPixels
        Picture1.ScaleMode = vbPixels
        Picture2.ScaleMode = vbPixels
        Picture3.ScaleMode = vbPixels
        Picture4.ScaleMode = vbPixels
        Picture5.ScaleMode = vbPixels
        Picture6.ScaleMode = vbPixels
         Printer.PaintPicture Picture3.Image, 3400, 100 ', wid, hgt
        Printer.PaintPicture Picture4.Image, 5000, 100 ', wid, hgt
        Printer.PaintPicture Picture2.Image, 3400, 310 ', wid, hgt
        Printer.PaintPicture Picture6.Image, 4100, 890 ', wid, hgt
        Printer.PaintPicture Picture1.Image, 3400, 1100 ', wid, hgt
        Printer.PaintPicture Picture5.Image, 3400, 1270 ', wid, hgt
       'wid = ScaleX(Picture2.ScaleWidth, Picture2.ScaleMode, _
        Printer.ScaleMode)
        'hgt = ScaleY(Picture2.ScaleHeight, _
        Picture2.ScaleMode, Printer.ScaleMode)
    '
    '    ' Draw the box.
        'Printer.Line (500, 500)-Step(wid, hgt), , B
        
        ' Finish printing.
        Printer.EndDoc
        i = i - 2
    Loop
    
    Exit Function
eRRhAND:
    MsgBox Err.Description
End Function


Private Function FILL_PREVIIOUSRATE()
    If LBLCUSTOMER.Caption = "" Then Exit Function
    Set GRDPRERATE.DataSource = Nothing

    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, SALES_TAX, REF_NO, CATEGORY, ITEM_COST, VCH_NO, LINE_NO, TRX_YEAR, MRP  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & LBLCUSTOMER.Caption & "'  ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, SALES_TAX, REF_NO, CATEGORY, ITEM_COST, VCH_NO, LINE_NO, TRX_YEAR, MRP  From TRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO') AND M_USER_ID = '" & LBLCUSTOMER.Caption & "' ORDER BY VCH_DATE DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If

    If PHY_PRERATE.RecordCount > 0 Then
        Fram.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        
        GRDPRERATE.Columns(0).Caption = "TYPE"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "BILL DATE"
        GRDPRERATE.Columns(4).Caption = "SOLD QTY"
        GRDPRERATE.Columns(5).Caption = "RATE"
        GRDPRERATE.Columns(6).Caption = "NET RATE"
        GRDPRERATE.Columns(7).Caption = "TAX"
        GRDPRERATE.Columns(8).Caption = "ARTICLE NO"
        GRDPRERATE.Columns(9).Caption = "CATEGORY"
        GRDPRERATE.Columns(10).Caption = "COST"
        GRDPRERATE.Columns(11).Caption = "VCH NO"
        GRDPRERATE.Columns(12).Caption = "LINE NO"
        GRDPRERATE.Columns(13).Caption = "YEAR"
        GRDPRERATE.Columns(14).Caption = "MRP"

        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 3500
        GRDPRERATE.Columns(3).Width = 1400
        GRDPRERATE.Columns(4).Width = 1100
        GRDPRERATE.Columns(5).Width = 1100
        GRDPRERATE.Columns(6).Width = 1100
        GRDPRERATE.Columns(7).Width = 1100
        GRDPRERATE.Columns(8).Width = 1100
        GRDPRERATE.Columns(9).Width = 1100
        GRDPRERATE.Columns(10).Width = 1100
        GRDPRERATE.Columns(11).Width = 0
        GRDPRERATE.Columns(12).Width = 0
        GRDPRERATE.Columns(13).Width = 0
        GRDPRERATE.Columns(14).Width = 1100

        GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).Text
        GRDPRERATE.SetFocus
    Else
        If MsgBox("This Item has not been sold Yet!! Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Else
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TxtSale_Rate.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        End If
    End If

End Function


Private Function FILL_PRERATEWITHBARCODE()
    If LBLCUSTOMER.Caption = "" Then Exit Function
    Set GRDPRERATE.DataSource = Nothing

    If PRERATE_FLAG = True Then
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, SALES_TAX, REF_NO, CATEGORY, ITEM_COST, VCH_NO, LINE_NO  From TRXFILE  WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO')  ORDER BY VCH_DATE DESC, TRX_TYPE DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    Else
        PHY_PRERATE.Close
        PHY_PRERATE.Open "Select TRX_TYPE, ITEM_CODE, ITEM_NAME, VCH_DATE, QTY, P_RETAILWOTAX, P_RETAIL, SALES_TAX, REF_NO, CATEGORY, ITEM_COST, VCH_NO, LINE_NO  From TRXFILE  WHERE BARCODE = '" & Trim(TxtBarcode.Text) & "' AND (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE = 'SI' OR TRX_TYPE = 'RI' OR TRX_TYPE = 'WO')  ORDER BY VCH_DATE DESC, TRX_TYPE DESC, VCH_NO DESC ", db, adOpenStatic, adLockReadOnly
        PRERATE_FLAG = False
    End If

    If PHY_PRERATE.RecordCount > 0 Then
        Fram.Enabled = False
        fRMEPRERATE.Visible = True
        Set GRDPRERATE.DataSource = PHY_PRERATE
        
        GRDPRERATE.Columns(0).Caption = "TYPE"
        GRDPRERATE.Columns(1).Caption = "ITEM CODE"
        GRDPRERATE.Columns(2).Caption = "ITEM NAME"
        GRDPRERATE.Columns(3).Caption = "BILL DATE"
        GRDPRERATE.Columns(4).Caption = "SOLD QTY"
        GRDPRERATE.Columns(5).Caption = "RATE"
        GRDPRERATE.Columns(6).Caption = "NET RATE"
        GRDPRERATE.Columns(7).Caption = "TAX"
        GRDPRERATE.Columns(8).Caption = "ARTICLE NO"
        GRDPRERATE.Columns(9).Caption = "CATEGORY"
        GRDPRERATE.Columns(10).Caption = "COST"
        GRDPRERATE.Columns(11).Caption = "VCH NO"
        GRDPRERATE.Columns(12).Caption = "LINE NO"

        GRDPRERATE.Columns(0).Visible = False
        GRDPRERATE.Columns(1).Visible = False
        GRDPRERATE.Columns(2).Width = 3500
        GRDPRERATE.Columns(3).Width = 1400
        GRDPRERATE.Columns(4).Width = 1200
        GRDPRERATE.Columns(5).Width = 1200
        GRDPRERATE.Columns(6).Width = 1200
        GRDPRERATE.Columns(7).Width = 1200
        GRDPRERATE.Columns(8).Width = 1200
        GRDPRERATE.Columns(9).Width = 1200
        GRDPRERATE.Columns(10).Width = 1200
        GRDPRERATE.Columns(11).Width = 0
        GRDPRERATE.Columns(12).Width = 0
       

        GRDPRERATE.SetFocus
        LBLHEAD(2).Caption = GRDPRERATE.Columns(2).Text
        GRDPRERATE.SetFocus
    Else
        If MsgBox("This Item has not been sold Yet!! Do You Want to Continue...?", vbYesNo, "SALES RETURN..") = vbYes Then
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Else
            Set GRDPRERATE.DataSource = Nothing
            fRMEPRERATE.Visible = False
            Fram.Enabled = True
            TxtSale_Rate.Enabled = False
            TXTPRODUCT.Enabled = True
            TXTPRODUCT.SetFocus
        End If
    End If

End Function



