VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMWARRANTY 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Return Items to the Supplier for Warranty Claim"
   ClientHeight    =   9585
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17010
   Icon            =   "Frmwarranty.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9585
   ScaleWidth      =   17010
   Begin MSMask.MaskEdBox TXTEXPIRY 
      Height          =   285
      Left            =   9570
      TabIndex        =   46
      Top             =   1695
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   330
      Picture         =   "Frmwarranty.frx":030A
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   44
      Top             =   30
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   0
      Picture         =   "Frmwarranty.frx":064C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   43
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   465
      Left            =   6225
      TabIndex        =   42
      Top             =   8250
      Visible         =   0   'False
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   820
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
   Begin VB.Frame FRMEGRDTMP 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   3915
      TabIndex        =   7
      Top             =   1935
      Visible         =   0   'False
      Width           =   5835
      Begin MSDataGridLib.DataGrid GRDPOPUP 
         Height          =   2535
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   4471
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
         Left            =   3135
         TabIndex        =   10
         Top             =   105
         Visible         =   0   'False
         Width           =   2610
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
         Left            =   90
         TabIndex        =   9
         Top             =   105
         Visible         =   0   'False
         Width           =   3045
      End
   End
   Begin VB.Frame FRMEITEM 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   3825
      TabIndex        =   5
      Top             =   1860
      Visible         =   0   'False
      Width           =   6030
      Begin MSDataGridLib.DataGrid GRDPOPUPITEM 
         Height          =   2835
         Left            =   105
         TabIndex        =   6
         Top             =   60
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5001
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
      Left            =   2370
      TabIndex        =   3
      Top             =   2475
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSDataListLib.DataCombo CMBSUPPLIERexp 
      Height          =   330
      Left            =   6900
      TabIndex        =   4
      Top             =   2685
      Visible         =   0   'False
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   255
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtactqty 
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
      Left            =   4470
      TabIndex        =   2
      Top             =   9075
      Width           =   1155
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   4725
      TabIndex        =   1
      Top             =   1605
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   9075
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
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
   Begin MSFlexGridLib.MSFlexGrid grdEXPIRYLIST 
      Height          =   7275
      Left            =   45
      TabIndex        =   0
      Top             =   30
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   12832
      _Version        =   393216
      Rows            =   1
      Cols            =   15
      FixedRows       =   0
      RowHeightMin    =   400
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
      HighLight       =   0
      FillStyle       =   1
      Appearance      =   0
      GridLineWidth   =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0FF&
      Height          =   1785
      Left            =   15
      TabIndex        =   11
      Top             =   7245
      Width           =   12255
      Begin VB.CheckBox CHKSELECT 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4350
         TabIndex        =   45
         Top             =   885
         Width           =   1320
      End
      Begin VB.TextBox TXTTRXTYPE 
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
         Left            =   6330
         TabIndex        =   27
         Top             =   1725
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
         Left            =   4500
         TabIndex        =   26
         Top             =   1695
         Visible         =   0   'False
         Width           =   690
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
         Left            =   2640
         TabIndex        =   25
         Top             =   1710
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
         Height          =   330
         Left            =   5160
         MaxLength       =   15
         TabIndex        =   24
         Top             =   360
         Width           =   1515
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
         Left            =   2655
         TabIndex        =   23
         Top             =   2085
         Visible         =   0   'False
         Width           =   690
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
         Height          =   390
         Left            =   7560
         TabIndex        =   22
         Top             =   735
         Width           =   1100
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
         Height          =   390
         Left            =   8760
         TabIndex        =   21
         Top             =   735
         Width           =   1100
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
         Height          =   390
         Left            =   11085
         TabIndex        =   20
         Top             =   735
         Width           =   1100
      End
      Begin VB.CommandButton CMDPRINT 
         Caption         =   "&PRINT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9915
         TabIndex        =   19
         Top             =   735
         Width           =   1100
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
         Height          =   330
         Left            =   4410
         MaxLength       =   7
         TabIndex        =   18
         Top             =   360
         Width           =   735
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
         Height          =   330
         Left            =   600
         TabIndex        =   17
         Top             =   360
         Width           =   3780
      End
      Begin VB.TextBox TXTSLNO 
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
         Height          =   330
         Left            =   45
         TabIndex        =   16
         Top             =   360
         Width           =   540
      End
      Begin VB.TextBox txtinvno 
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
         Left            =   9795
         MaxLength       =   15
         TabIndex        =   15
         Top             =   375
         Width           =   1050
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         Height          =   480
         Left            =   60
         TabIndex        =   12
         Top             =   630
         Width           =   4230
         Begin VB.OptionButton optnamesort 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Sort by Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   45
            TabIndex        =   14
            Top             =   180
            Width           =   1830
         End
         Begin VB.OptionButton optsuplsort 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Sort by Supplier"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   13
            Top             =   165
            Value           =   -1  'True
            Width           =   2145
         End
      End
      Begin MSDataListLib.DataCombo CMBSUPPLIER 
         Height          =   330
         Left            =   6690
         TabIndex        =   28
         Top             =   360
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         ForeColor       =   255
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSMask.MaskEdBox txtinvdate 
         Height          =   315
         Left            =   10860
         TabIndex        =   29
         Top             =   375
         Width           =   1320
         _ExtentX        =   2328
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
         Left            =   3360
         TabIndex        =   41
         Top             =   1710
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   20
         Left            =   7065
         TabIndex        =   40
         Top             =   1785
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "TRX TYPE."
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
         Index           =   19
         Left            =   5190
         TabIndex        =   39
         Top             =   1740
         Visible         =   0   'False
         Width           =   1080
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
         Left            =   1500
         TabIndex        =   38
         Top             =   1710
         Visible         =   0   'False
         Width           =   1080
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   7
         Left            =   5160
         TabIndex        =   37
         Top             =   135
         Width           =   1515
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
         Left            =   1515
         TabIndex        =   36
         Top             =   2100
         Visible         =   0   'False
         Width           =   1080
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   10
         Left            =   4410
         TabIndex        =   35
         Top             =   135
         Width           =   735
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
         ForeColor       =   &H0000FFFF&
         Height          =   240
         Index           =   9
         Left            =   600
         TabIndex        =   34
         Top             =   135
         Width           =   3780
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
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   8
         Left            =   45
         TabIndex        =   33
         Top             =   135
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "Supplier"
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
         Height          =   225
         Index           =   2
         Left            =   6690
         TabIndex        =   32
         Top             =   135
         Width           =   3090
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "INV #"
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
         Index           =   3
         Left            =   9795
         TabIndex        =   31
         Top             =   135
         Width           =   1050
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "INV Date"
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
         Index           =   4
         Left            =   10860
         TabIndex        =   30
         Top             =   135
         Width           =   1320
      End
   End
End
Attribute VB_Name = "FRMWARRANTY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim TMPREC As New ADODB.Recordset
Dim TMPFLAG As Boolean

Dim PHY_BATCH As New ADODB.Recordset
Dim BATCH_FLAG As Boolean
Dim PHY_ITEM As New ADODB.Recordset
Dim ITEM_FLAG As Boolean
Dim ACT_REC As New ADODB.Recordset
Dim MFG_REC As New ADODB.Recordset

Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim NONSTOCK As Boolean
'Dim strChecked As String

Private Sub CHKSELECT_Click()
    Dim i As Long
    If grdEXPIRYLIST.Rows = 1 Then Exit Sub
    For i = 1 To grdEXPIRYLIST.Rows - 1
        If CHKSELECT.value = 1 Then
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picUnchecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 13) = "Y"
            End With
        Else
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
              .TextMatrix(.Row, 13) = "N"
            End With
        End If
    Next i
    Call fillcount
End Sub

Private Sub CMBSUPPLIER_GotFocus()
    CMBSUPPLIER.SelStart = 0
    CMBSUPPLIER.SelLength = Len(CMBSUPPLIER.Text)
End Sub

Private Sub CMBSUPPLIER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMBSUPPLIER.MatchedWithList = False Then
                MsgBox "Select Supplier from the list", vbOKOnly, "Expiry!!!"
                CMBSUPPLIER.SelStart = 0
                CMBSUPPLIER.SelLength = Len(CMBSUPPLIER.Text)
                CMBSUPPLIER.SetFocus
                Exit Sub
            End If
            CMBSUPPLIER.Enabled = False
            txtinvno.Enabled = True
            txtinvno.SetFocus
        Case vbKeyEscape
            CMBSUPPLIER.Enabled = False
            
    End Select
End Sub

Private Sub CMBSUPPLIER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMBSUPPLIERexp_Click(Area As Integer)
    CMBSUPPLIERexp.SelStart = 0
    CMBSUPPLIERexp.SelLength = Len(CMBSUPPLIERexp.Text)
End Sub

Private Sub CMBSUPPLIERexp_KeyDown(KeyCode As Integer, Shift As Integer)
       Dim RSTSUPPLIER As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(CMBSUPPLIERexp.Text) = "" Then Exit Sub
            If CMBSUPPLIERexp.MatchedWithList = False Then
                MsgBox "Select Supplier from the list", vbOKOnly, "Expiry!!!"
                CMBSUPPLIERexp.SelStart = 0
                CMBSUPPLIERexp.SelLength = Len(CMBSUPPLIER.Text)
                CMBSUPPLIERexp.SetFocus
                Exit Sub
            End If
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                RSTSUPPLIER!DIST_NAME = Trim(CMBSUPPLIERexp.Text)
                RSTSUPPLIER.Update
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
            
            grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = Trim(CMBSUPPLIERexp.Text)
            grdEXPIRYLIST.Enabled = True
            CMBSUPPLIERexp.Visible = False
            grdEXPIRYLIST.SetFocus
        Case vbKeyEscape
            CMBSUPPLIERexp.Visible = False
            grdEXPIRYLIST.SetFocus
    End Select
End Sub

Private Sub CMBSUPPLIERexp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMBSUPPLIERexp_LostFocus()
    CMBSUPPLIERexp.Visible = False
End Sub

Private Sub CMDADD_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    Dim FCODE1, FCODE2, FCODE3, FCODE4, FCODE5, FCODE6, FCODE7, FCODE10, FCODE12, FCODE14
    Dim FCODE, FCODE8, FCODE9, FCODE11, FCODE13, FCODE15   As Long

    On Error GoTo ErrHand
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select MAX(EX_SLNO) from WAR_LIST", db, adOpenForwardOnly
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        If IsNull(RSTTRXFILE.Fields(0)) Then
             FCODE = 1
        Else
            FCODE = Val(RSTTRXFILE.Fields(0)) + 1
        End If
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    db.Execute ("Insert into WAR_LIST values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "','" & FCODE7 & "','" & FCODE8 & "','" & FCODE9 & "','" & FCODE10 & "','" & FCODE11 & "','" & FCODE12 & "','" & FCODE13 & "','" & FCODE14 & "' )")
        
    grdEXPIRYLIST.Rows = grdEXPIRYLIST.Rows + 1
    grdEXPIRYLIST.FixedRows = 1
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 1) = FCODE1
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 2) = FCODE4
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 3) = FCODE5
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 4) = FCODE13
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 5) = FCODE8
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 6) = FCODE6
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 7) = FCODE7
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 8) = FCODE9
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 9) = FCODE
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 10) = FCODE2 'invoice no
    grdEXPIRYLIST.TextMatrix(Val(TXTSLNO.Text), 11) = FCODE3
    
    If NONSTOCK = True Then GoTo SKIP
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            !ISSUE_QTY = !ISSUE_QTY + FCODE15
            !CLOSE_QTY = !CLOSE_QTY - FCODE15
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM RTRXFILE WHERE RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        If Not (.EOF And .BOF) Then
            If (IsNull(!ISSUE_QTY)) Then !ISSUE_QTY = 0
            !ISSUE_QTY = !ISSUE_QTY + FCODE15
            
            If (IsNull(!BAL_QTY)) Then !BAL_QTY = 0
            !BAL_QTY = !BAL_QTY - FCODE15
            
            RSTTRXFILE.Update
        End If
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing

SKIP:
    
    TXTSLNO.Text = grdEXPIRYLIST.Rows
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTTRXTYPE.Text = ""
    
    TXTINVDATE.Text = "  /  /    "
    TXTQTY.Text = ""
    TxtActqty.Text = ""
    txtinvno.Text = ""
    CMBSUPPLIER.Text = ""
    txtBatch.Text = ""
    cmdadd.Enabled = False
    M_EDIT = False
    TXTPRODUCT.Enabled = True
    TXTPRODUCT.SetFocus
    grdEXPIRYLIST.TopRow = grdEXPIRYLIST.Rows - 1
    NONSTOCK = False
Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    If grdEXPIRYLIST.Rows = 1 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE SERIAL NO. " & """" & grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 0) & """" & " FROM THE LIST", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
 
    db.Execute ("Delete from WAR_LIST where WAR_LIST.EX_SLNO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 9)) & " ")
    
    Call FILLEXPIRYGRID
    
    TXTSLNO.Text = Val(grdEXPIRYLIST.Rows)
    TXTPRODUCT.Text = ""
    TXTITEMCODE.Text = ""
    TXTVCHNO.Text = ""
    TXTLINENO.Text = ""
    TXTINVDATE.Text = "  /  /    "
    TXTTRXTYPE.Text = ""
    TXTQTY.Text = ""
    TxtActqty.Text = ""
    txtinvno.Text = ""
    CMBSUPPLIER.Text = ""
    txtBatch.Text = ""
    cmdadd.Enabled = False
    M_EDIT = False
    
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub


Private Sub CmdPrint_Click()
    If grdEXPIRYLIST.Rows = 1 Then Exit Sub
    MDIMAIN.Enabled = False
    Enabled = False
    FRMPRINTEXP.Visible = True
End Sub

Private Sub Form_Activate()
    If TXTPRODUCT.Enabled = True Then TXTPRODUCT.SetFocus
    If TXTQTY.Enabled = True Then TXTQTY.SetFocus
    If txtBatch.Enabled = True Then txtBatch.SetFocus
    If cmdadd.Enabled = True Then cmdadd.SetFocus
End Sub

Private Sub Form_Load()
    Dim rstfillcombo As ADODB.Recordset
    On Error GoTo ErrHand
    
    PHYFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    
    Set CMBSUPPLIER.DataSource = Nothing
    Set CMBSUPPLIERexp.DataSource = Nothing
    ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenForwardOnly
    Set CMBSUPPLIER.RowSource = ACT_REC
    Set CMBSUPPLIERexp.RowSource = ACT_REC
    CMBSUPPLIER.ListField = "ACT_NAME"
    CMBSUPPLIER.BoundColumn = "ACT_CODE"
    CMBSUPPLIERexp.ListField = "ACT_NAME"
    CMBSUPPLIERexp.BoundColumn = "ACT_CODE"
    
    Call Fillgrid
    
    Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHYFLAG = False Then PHY.Close
    If TMPFLAG = False Then TMPREC.Close
    If BATCH_FLAG = False Then PHY_BATCH.Close
    If ITEM_FLAG = False Then PHY_ITEM.Close
    ACT_REC.Close
    'MFG_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub grdEXPIRYLIST_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If grdEXPIRYLIST.Rows = 1 Then Exit Sub
    If grdEXPIRYLIST.Col <> 1 Then Exit Sub
    With grdEXPIRYLIST
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If grdEXPIRYLIST.Col = 0 Then
                If grdEXPIRYLIST.CellPicture = picChecked Then
                    Set grdEXPIRYLIST.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 13) = "Y"
                    Call fillcount
                Else
                    Set grdEXPIRYLIST.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 13) = "N"
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Sub grdEXPIRYLIST_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeySpace
            Call grdEXPIRYLIST_Click
        Case vbKeyReturn
            Select Case grdEXPIRYLIST.Col
                Case 5, 6, 9
                    Select Case grdEXPIRYLIST.Col
                        Case 9 'bILL nO
                            TXTsample.MaxLength = 15
                        Case 5 'qty
                            TXTsample.MaxLength = 6
                        Case 6 'batch
                            TXTsample.MaxLength = 30
                    End Select
                    TXTsample.Visible = True
                    TXTsample.Top = grdEXPIRYLIST.CellTop + 125
                    TXTsample.Left = grdEXPIRYLIST.CellLeft + 25
                    TXTsample.Width = grdEXPIRYLIST.CellWidth + 50
                    TXTsample.Text = grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col)
                    TXTsample.SetFocus
        
                Case 8
                    CMBSUPPLIERexp.Visible = True
                    CMBSUPPLIERexp.Top = grdEXPIRYLIST.CellTop + 125
                    CMBSUPPLIERexp.Left = grdEXPIRYLIST.CellLeft + 25
                   ' CMBSUPPLIER.Width = grdEXPIRYLIST.CellWidth
                    CMBSUPPLIERexp.Text = grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col)
                    CMBSUPPLIERexp.SetFocus
                Case 10
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = grdEXPIRYLIST.CellTop + 325
                    TXTEXPIRY.Left = grdEXPIRYLIST.CellLeft + 75
                    TXTEXPIRY.Width = grdEXPIRYLIST.CellWidth
                    TXTEXPIRY.Height = grdEXPIRYLIST.CellHeight
                    If Not (IsDate(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col))) Then
                        TXTEXPIRY.Text = "  /  /    "
                    Else
                        TXTEXPIRY.Text = Format(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col), "DD/MM/YYYY")
                    End If
                    
                    TXTEXPIRY.SetFocus
            End Select
    End Select
End Sub

Private Sub GRDPOPUP_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            TxtActqty.Text = GRDPOPUP.Columns(1)
            TXTQTY.Text = GRDPOPUP.Columns(1)
            txtBatch.Text = GRDPOPUP.Columns(0)
            
            TXTVCHNO.Text = GRDPOPUP.Columns(7)
            TXTLINENO.Text = GRDPOPUP.Columns(8)
            TXTTRXTYPE.Text = GRDPOPUP.Columns(9)
            
            txtinvno.Text = GRDPOPUP.Columns(12)
            TXTINVDATE.Text = IIf(IsNull(GRDPOPUP.Columns(14)), "  /  /    ", GRDPOPUP.Columns(14))
            
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT ACT_NAME  FROM ACTMAST WHERE ACT_CODE = '" & GRDPOPUP.Columns(13) & "'", db, adOpenForwardOnly
            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                CMBSUPPLIER.Text = RSTTRXFILE!ACT_NAME
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenForwardOnly
'            If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                CMBMFGR.BoundText = Trim(RSTTRXFILE!MANUFACTURER)
'            End If
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTQTY.Text = ""
            TxtActqty.Text = ""
            txtinvno.Text = ""
            CMBSUPPLIER.Text = ""
            TXTVCHNO.Text = ""
            TXTINVDATE.Text = "  /  /    "
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            
            Set GRDPOPUP.DataSource = Nothing
            FRMEGRDTMP.Visible = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
        
    End Select
End Sub

Private Sub GRDPOPUPITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            M_STOCK = Val(GRDPOPUPITEM.Columns(2))
            TXTPRODUCT.Text = GRDPOPUPITEM.Columns(1)
            TXTITEMCODE.Text = GRDPOPUPITEM.Columns(0)
            If M_STOCK = 0 Then
                If MsgBox("NO STOCK AVAILABLE!!! DO YOU WANT TO CONTINUE....", vbYesNo, "EXPIRY!!!") = vbYes Then
                    FRMEITEM.Visible = False
                    NONSTOCK = True
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                End If
                Exit Sub
            End If
'            For i = 1 To grdEXPIRYLIST.Rows - 1
'                If Trim(grdEXPIRYLIST.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
'                    If MsgBox("This Item Already exists.... Do yo want to add this item", vbYesNo, "BILL..") = vbNo Then
'                        Set GRDPOPUPITEM.DataSource = Nothing
'                        FRMEITEM.Visible = False
'                        FRMEMAIN.Enabled = True
'                        TXTPRODUCT.Enabled = True
'                        TXTQTY.Enabled = False
'                        TXTPRODUCT.SetFocus
'                        Exit Sub
'                    Else
'                        Exit For
'                    End If
'                End If
'            Next i
            Set GRDPOPUPITEM.DataSource = Nothing
            If ITEM_FLAG = True Then
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
                ITEM_FLAG = False
            Else
                PHY_ITEM.Close
                PHY_ITEM.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
                ITEM_FLAG = False
            End If
            Set GRDPOPUPITEM.DataSource = PHY_ITEM
            If PHY_ITEM.RecordCount = 1 Then
                TXTQTY.Text = GRDPOPUPITEM.Columns(2)
                TxtActqty.Text = GRDPOPUPITEM.Columns(2)
                txtBatch.Text = GRDPOPUPITEM.Columns(6)
                
                TXTVCHNO.Text = GRDPOPUPITEM.Columns(8)
                TXTLINENO.Text = GRDPOPUPITEM.Columns(9)
                TXTTRXTYPE.Text = GRDPOPUPITEM.Columns(10)
            
                txtinvno.Text = GRDPOPUPITEM.Columns(12)
                TXTINVDATE.Text = IIf(IsNull(GRDPOPUPITEM.Columns(14)), "  /  /    ", GRDPOPUPITEM.Columns(14))
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "SELECT ACT_NAME  FROM ACTMAST WHERE ACT_CODE = '" & GRDPOPUPITEM.Columns(13) & "'", db, adOpenForwardOnly
                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                    CMBSUPPLIER.Text = RSTTRXFILE!ACT_NAME
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
'                Set RSTTRXFILE = New ADODB.Recordset
'                RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenForwardOnly
'                If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                    CMBMFGR.BoundText = Trim(RSTTRXFILE!MANUFACTURER)
'                End If
'                RSTTRXFILE.Close
'                Set RSTTRXFILE = Nothing
                
                
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEITEM.Visible = False
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY_ITEM.RecordCount > 1 Then
                Set GRDPOPUPITEM.DataSource = Nothing
                FRMEGRDTMP.Visible = False
                Call FILL_BATCHGRID
            End If
        Case vbKeyEscape
            TXTQTY.Text = ""
            TxtActqty.Text = ""
            txtinvno.Text = ""
            CMBSUPPLIER.Text = ""
            TXTVCHNO.Text = ""
            TXTINVDATE.Text = "  /  /    "
            TXTLINENO.Text = ""
            TXTTRXTYPE.Text = ""
            Set GRDPOPUPITEM.DataSource = Nothing
            FRMEITEM.Visible = False
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            TXTPRODUCT.SetFocus
            
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub optnamesort_Click()
    FILLEXPIRYGRID
End Sub

Private Sub optsuplsort_Click()
    FILLEXPIRYGRID
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
        Case vbKeyEscape
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
            'txtexpdate.SetFocus
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

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.Visible = False
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTINVDATE.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTINVDATE.Enabled = False
            txtinvno.Enabled = True
            txtinvno.SetFocus
    End Select
End Sub

Private Sub txtinvno_GotFocus()
    txtinvno.SelStart = 0
    txtinvno.SelLength = Len(txtinvno.Text)
End Sub

Private Sub txtinvno_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtinvno.Enabled = False
            TXTINVDATE.Enabled = True
            TXTINVDATE.SetFocus
        Case vbKeyEscape
            txtinvno.Enabled = False
            CMBSUPPLIER.Enabled = True
            CMBSUPPLIER.SetFocus
    End Select
End Sub

Private Sub txtinvno_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
    NONSTOCK = False
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTPRODUCT.Text) = "" Then Exit Sub
            If NONSTOCK = True Then GoTo SKIP
            TXTINVDATE.Text = "  /  /    "
            TXTQTY.Text = ""
            TxtActqty.Text = ""
            txtinvno.Text = ""
            CMBSUPPLIER.Text = ""
            txtBatch.Text = ""
            'If Len(TXTPRODUCT.Text) < 2 Then Exit Sub
            'Call STOCKCORRECTION
            
            Set grdtmp.DataSource = Nothing
            If PHYFLAG = True Then
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME,CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
                PHYFLAG = False
            Else
                PHY.Close
                PHY.Open "Select DISTINCT ITEM_CODE,ITEM_NAME,CLOSE_QTY From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
                PHYFLAG = False
            End If
            
            Set grdtmp.DataSource = PHY
            If PHY.RecordCount = 1 Then
                TXTITEMCODE.Text = grdtmp.Columns(0)
                TXTPRODUCT.Text = grdtmp.Columns(1)
'                For i = 1 To grdEXPIRYLIST.Rows - 1
'                    If Trim(grdEXPIRYLIST.TextMatrix(i, 12)) = Trim(TXTITEMCODE.Text) Then
'                        If MsgBox("This Item Already exists... Do yo want to add this item again", vbYesNo, "BILL..") = vbNo Then
'                            Exit Sub
'                        Else
'                            Exit For
'                        End If
'                    End If
'                Next i
                Set grdtmp.DataSource = Nothing
                If TMPFLAG = True Then
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
                    TMPFLAG = False
                Else
                    TMPREC.Close
                    TMPREC.Open "Select ITEM_CODE, ITEM_NAME, BAL_QTY, MRP, SALES_TAX, UNIT, REF_NO, EXP_DATE, VCH_NO, LINE_NO, TRX_TYPE, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
                    TMPFLAG = False
                End If
                Set grdtmp.DataSource = TMPREC
                If TMPREC.RecordCount = 1 Then
                    TXTQTY.Text = grdtmp.Columns(2)
                    TxtActqty.Text = grdtmp.Columns(2)
                    txtBatch.Text = grdtmp.Columns(6)
                    txtinvno.Text = grdtmp.Columns(12)
                    TXTVCHNO.Text = grdtmp.Columns(8)
                    TXTLINENO.Text = grdtmp.Columns(9)
                    TXTTRXTYPE.Text = grdtmp.Columns(10)
                    TXTINVDATE.Text = IIf(IsNull(grdtmp.Columns(14)), "  /  /    ", grdtmp.Columns(14))
                    Set RSTTRXFILE = New ADODB.Recordset
                    RSTTRXFILE.Open "SELECT ACT_NAME  FROM ACTMAST WHERE ACT_CODE = '" & grdtmp.Columns(13) & "'", db, adOpenForwardOnly
                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                        CMBSUPPLIER.Text = RSTTRXFILE!ACT_NAME
                    End If
                    RSTTRXFILE.Close
                    Set RSTTRXFILE = Nothing
                    
'                    Set RSTTRXFILE = New ADODB.Recordset
'                    RSTTRXFILE.Open "SELECT MANUFACTURER  FROM ITEMMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenForwardOnly
'                    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'                        CMBMFGR.BoundText = Trim(RSTTRXFILE!MANUFACTURER)
'                    End If
'                    RSTTRXFILE.Close
'                    Set RSTTRXFILE = Nothing
            
                    
                    TXTPRODUCT.Enabled = False
                    TXTQTY.Enabled = True
                    TXTQTY.SetFocus
                    
                    Exit Sub
                ElseIf TMPREC.RecordCount = 0 Then
                    If MsgBox("NO STOCK AVAILABLE!!! DO YOU WANT TO CONTINUE....", vbYesNo, "EXPIRY!!!") = vbNo Then
                        TXTQTY.Text = 0
                        TxtActqty.Text = 0
                        TXTPRODUCT.Enabled = True
                        TXTQTY.Enabled = False
                        TXTPRODUCT.SelStart = 0
                        TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
                        TXTPRODUCT.SetFocus
                    Else
                        NONSTOCK = True
                        TXTPRODUCT.Enabled = False
                        TXTQTY.Enabled = True
                        TXTQTY.SetFocus
                    End If
                    Exit Sub
                ElseIf TMPREC.RecordCount > 1 Then
                    Call FILL_BATCHGRID
                    Exit Sub
                End If
SKIP:
                TXTPRODUCT.Enabled = False
                TXTQTY.Enabled = True
                txtBatch.Enabled = False
                TXTQTY.SetFocus
                Exit Sub
            ElseIf PHY.RecordCount > 1 Then
                'FRMSUB.grdsub.Columns(0).Visible = True
                'FRMSUB.grdsub.Columns(1).Caption = "ITEM NAME"
                'FRMSUB.grdsub.Columns(1).Width = 3200
                'FRMSUB.grdsub.Columns(2).Caption = "QTY"
                'FRMSUB.grdsub.Columns(2).Width = 1300
                Call FILL_ITEMGRID
            ElseIf PHY.RecordCount = 0 Then
                MsgBox "No Such Item Exists!!!!!", vbOKOnly, "BILL.."
            End If
        
            TXTPRODUCT.Enabled = True
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
        Case vbKeyF5
            NONSTOCK = True
            '''lblnonstockdisplay.Visible = True
        Case vbKeyEscape
            CmdExit.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
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
    Dim RSTTRXFILE As ADODB.Recordset
    Dim i As Long
    
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then Exit Sub
            If NONSTOCK = True Then GoTo SKIP
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "SELECT BAL_QTY  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenForwardOnly
             If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                If (IsNull(RSTTRXFILE!BAL_QTY)) Then RSTTRXFILE!BAL_QTY = 0
                i = RSTTRXFILE!BAL_QTY
            End If
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
'            txtexpdate.Text = Date
'            Set RSTTRXFILE = New ADODB.Recordset
'            RSTTRXFILE.Open "SELECT EXP_DATE  FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'AND RTRXFILE.TRX_TYPE = '" & Trim(TXTTRXTYPE.Text) & "' AND RTRXFILE.VCH_NO = " & Val(TXTVCHNO.Text) & " AND RTRXFILE.LINE_NO = " & Val(TXTLINENO.Text) & "", db, adOpenForwardOnly
'            If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
'                If (IsNull(RSTTRXFILE!EXP_DATE)) Then
'                    txtexpdate.Text = Date
'                Else
'                    txtexpdate.Text = RSTTRXFILE!EXP_DATE
'                End If
'            End If
'            RSTTRXFILE.Close
'            Set RSTTRXFILE = Nothing
            If Val(TXTQTY.Text) > i Then
                If MsgBox("Available Stock is " & i & " Do You want to Continue..?", vbYesNo, "EXPIRY!!!") = vbNo Then
                    'MsgBox "Available Stock is " & i, vbOKOnly, "BILL.."
                    TXTQTY.SelStart = 0
                    TXTQTY.SelLength = Len(TXTQTY.Text)
                    Exit Sub
                End If
            End If
SKIP:
            TXTPRODUCT.Enabled = False
            TXTQTY.Enabled = False
            txtBatch.Enabled = False
         Case vbKeyEscape
            TXTQTY.Enabled = False
            TXTQTY.Enabled = True
            TXTQTY.SetFocus
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

'Private Sub TXTDISCAMOUNT_KeyPress(KeyAscii As Integer)
'    Select Case KeyAscii
'        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
'    End Select
'End Sub


Function FILL_BATCHGRID()

    FRMEGRDTMP.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEITEM.Visible = False
    
    If BATCH_FLAG = True Then
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
        BATCH_FLAG = False
    Else
        PHY_BATCH.Close
        PHY_BATCH.Open "Select REF_NO, BAL_QTY, EXP_DATE, MRP, SALES_TAX,  ITEM_CODE, ITEM_NAME, VCH_NO, LINE_NO, TRX_TYPE, UNIT, ITEM_COST, PINV, M_USER_ID, VCH_DATE From RTRXFILE  WHERE ITEM_CODE = '" & TXTITEMCODE.Text & "' AND BAL_QTY > 0 ORDER BY EXP_DATE", db, adOpenForwardOnly
        BATCH_FLAG = False
    End If
    
    Set GRDPOPUP.DataSource = PHY_BATCH
    GRDPOPUP.Columns(0).Caption = "BATCH NO."
    GRDPOPUP.Columns(1).Caption = "QTY"
    GRDPOPUP.Columns(2).Caption = "EXP DATE"
    GRDPOPUP.Columns(3).Caption = "PRICE"
    GRDPOPUP.Columns(4).Caption = "TAX"
    GRDPOPUP.Columns(5).Caption = "Item Code"
    GRDPOPUP.Columns(6).Caption = "Item Name"
    GRDPOPUP.Columns(7).Caption = "VCH No"
    GRDPOPUP.Columns(8).Caption = "Line No"
    GRDPOPUP.Columns(9).Caption = "Trx Type"
    GRDPOPUP.Columns(10).Caption = "UNIT"
    GRDPOPUP.Columns(11).Caption = "COST"
    GRDPOPUP.Columns(12).Caption = "PINV"
    GRDPOPUP.Columns(13).Caption = "ACT_CODE"
    GRDPOPUP.Columns(14).Caption = "INV DATE"
    
    GRDPOPUP.Columns(0).Width = 1400
    GRDPOPUP.Columns(1).Width = 900
    GRDPOPUP.Columns(2).Width = 1400
    GRDPOPUP.Columns(3).Width = 1000
    GRDPOPUP.Columns(4).Width = 900
    
    GRDPOPUP.Columns(5).Visible = False
    GRDPOPUP.Columns(6).Visible = False
    GRDPOPUP.Columns(7).Visible = False
    GRDPOPUP.Columns(8).Visible = False
    GRDPOPUP.Columns(9).Visible = False
    GRDPOPUP.Columns(10).Visible = False
    GRDPOPUP.Columns(11).Visible = False
    GRDPOPUP.Columns(12).Visible = False
    
    GRDPOPUP.SetFocus
    LBLHEAD(0).Caption = GRDPOPUP.Columns(6).Text
    LBLHEAD(9).Visible = True
    LBLHEAD(0).Visible = True
End Function

Function FILL_ITEMGRID()
    FRMEITEM.Visible = True
    Set GRDPOPUP.DataSource = Nothing
    Set GRDPOPUPITEM.DataSource = Nothing
    FRMEGRDTMP.Visible = False
    
    
    If ITEM_FLAG = True Then
        PHY_ITEM.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY,MRP From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
        ITEM_FLAG = False
    Else
        PHY_ITEM.Close
        PHY_ITEM.Open "Select DISTINCT ITEM_CODE,ITEM_NAME, CLOSE_QTY,MRP From ITEMMAST  WHERE ITEM_NAME Like '%" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenForwardOnly
        ITEM_FLAG = False
    End If

    Set GRDPOPUPITEM.DataSource = PHY_ITEM
    'GRDPOPUPITEM.RowHeight = 250
    GRDPOPUPITEM.Columns(0).Visible = False
    GRDPOPUPITEM.Columns(1).Caption = "ITEM NAME"
    GRDPOPUPITEM.Columns(1).Width = 4500
    GRDPOPUPITEM.Columns(2).Caption = "QTY"
    GRDPOPUPITEM.Columns(3).Caption = "RATE"
    GRDPOPUPITEM.Columns(2).Width = 900
    GRDPOPUPITEM.Columns(3).Width = 900
    GRDPOPUPITEM.SetFocus
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTEXP As ADODB.Recordset
    Dim M_STOCK As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdEXPIRYLIST.Col
'                Case 4 ' ITEM NAME
'                    If Trim(TXTsample.Text) = "" Then Exit Sub
'                    Set RSTEXP = New ADODB.Recordset
'                    RSTEXP.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (RSTEXP.EOF And RSTEXP.BOF) Then
'                        RSTEXP!EX_ITEM = Trim(TXTsample.Text)
'                        RSTEXP.Update
'                    End If
'                    RSTEXP.Close
'                    Set RSTEXP = Nothing
'
'                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
'                    grdEXPIRYLIST.Enabled = True
'                    TXTsample.Visible = False
'                    grdEXPIRYLIST.SetFocus
'                Case 2  ' MFGR
'                    Set RSTEXP = New ADODB.Recordset
'                    RSTEXP.Open "SELECT * from WAR_TRXFILE where WAR_TRXFILE.EX_SLNO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 9)) & " ", DB, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (RSTEXP.EOF And RSTEXP.BOF) Then
'                        RSTEXP!EX_MFGR = Trim(TXTsample.Text)
'                        RSTEXP.Update
'                    End If
'                    RSTEXP.Close
'                     Set RSTEXP = Nothing
'
'                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
'                    grdEXPIRYLIST.Enabled = True
'                    TXTsample.Visible = False
'                    grdEXPIRYLIST.SetFocus
                Case 5   'QTY
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set RSTEXP = New ADODB.Recordset
                    RSTEXP.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTEXP.EOF And RSTEXP.BOF) Then
                        RSTEXP!QTY = Val(TXTsample.Text)
                        RSTEXP.Update
                    End If
                    RSTEXP.Close
                    Set RSTEXP = Nothing
                     
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
                    grdEXPIRYLIST.Enabled = True
                    TXTsample.Visible = False
                    grdEXPIRYLIST.SetFocus
                Case 6   'batch
                    Set RSTEXP = New ADODB.Recordset
                    RSTEXP.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTEXP.EOF And RSTEXP.BOF) Then
                        RSTEXP!REF_NO = Trim(TXTsample.Text)
                        RSTEXP.Update
                    End If
                    RSTEXP.Close
                    Set RSTEXP = Nothing
                     
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
                    grdEXPIRYLIST.Enabled = True
                    TXTsample.Visible = False
                    grdEXPIRYLIST.SetFocus
                Case 9  'INV NO
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set RSTEXP = New ADODB.Recordset
                    RSTEXP.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTEXP.EOF And RSTEXP.BOF) Then
                        RSTEXP!BILL_NO = Trim(TXTsample.Text)
                        RSTEXP.Update
                    End If
                    RSTEXP.Close
                    Set RSTEXP = Nothing
                    
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
                    grdEXPIRYLIST.Enabled = True
                    TXTsample.Visible = False
                    grdEXPIRYLIST.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdEXPIRYLIST.SetFocus
    End Select
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdEXPIRYLIST.Col
        Case 5, 4
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 1, 2, 6
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 8
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

Private Sub grdEXPIRYLIST_Scroll()
    TXTsample.Visible = False
    CMBSUPPLIERexp.Visible = False
    grdEXPIRYLIST.SetFocus
End Sub

Public Sub FILLEXPIRYGRID()
    Dim i As Long
    Dim rstrefresh As ADODB.Recordset
    
    i = 0
    grdEXPIRYLIST.TextMatrix(0, 0) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 1) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 2) = "MFGR"
    grdEXPIRYLIST.TextMatrix(0, 3) = "SUPPLIER"
    grdEXPIRYLIST.TextMatrix(0, 4) = "PACK"
    grdEXPIRYLIST.TextMatrix(0, 5) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 6) = "BATCH"
    grdEXPIRYLIST.TextMatrix(0, 7) = "EXPIRY"
    grdEXPIRYLIST.TextMatrix(0, 8) = "MRP"
    grdEXPIRYLIST.TextMatrix(0, 9) = "NO"
    grdEXPIRYLIST.TextMatrix(0, 10) = "INV NO"
    grdEXPIRYLIST.TextMatrix(0, 11) = "INV DATE"
    
    
    grdEXPIRYLIST.ColWidth(0) = 500
    grdEXPIRYLIST.ColWidth(1) = 2000
    grdEXPIRYLIST.ColWidth(2) = 1000
    grdEXPIRYLIST.ColWidth(3) = 2000
    grdEXPIRYLIST.ColWidth(4) = 1200
    grdEXPIRYLIST.ColWidth(5) = 800
    grdEXPIRYLIST.ColWidth(6) = 900
    grdEXPIRYLIST.ColWidth(7) = 1000
     grdEXPIRYLIST.ColWidth(8) = 900
    grdEXPIRYLIST.ColWidth(9) = 0
    
    grdEXPIRYLIST.ColAlignment(0) = 1
    grdEXPIRYLIST.ColAlignment(1) = 1
    grdEXPIRYLIST.ColAlignment(2) = 1
    grdEXPIRYLIST.ColAlignment(3) = 1
    grdEXPIRYLIST.ColAlignment(4) = 3
    grdEXPIRYLIST.ColAlignment(5) = 3
    grdEXPIRYLIST.ColAlignment(6) = 1
    grdEXPIRYLIST.ColAlignment(7) = 3
    On Error GoTo ErrHand
    i = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.Rows = 1
    Set rstrefresh = New ADODB.Recordset
    If optsuplsort.value = True Then
        rstrefresh.Open "SELECT * from WAR_LIST WHERE EX_FLAG ='N' ORDER BY EX_DISTI, EX_ITEM", db, adOpenStatic, adLockOptimistic, adCmdText
    Else
        rstrefresh.Open "SELECT * from WAR_LIST WHERE EX_FLAG ='N' ORDER BY EX_ITEM, EX_DISTI", db, adOpenStatic, adLockOptimistic, adCmdText
    End If
    Do Until rstrefresh.EOF
        i = i + 1
        grdEXPIRYLIST.Rows = grdEXPIRYLIST.Rows + 1
        grdEXPIRYLIST.FixedRows = 1
        grdEXPIRYLIST.TextMatrix(i, 0) = i
        grdEXPIRYLIST.TextMatrix(i, 1) = IIf(IsNull(rstrefresh!EX_ITEM), "", rstrefresh!EX_ITEM)
        grdEXPIRYLIST.TextMatrix(i, 2) = IIf(IsNull(rstrefresh!EX_MFGR), "", rstrefresh!EX_MFGR)
        grdEXPIRYLIST.TextMatrix(i, 3) = IIf(IsNull(rstrefresh!EX_DISTI), "", rstrefresh!EX_DISTI)
        grdEXPIRYLIST.TextMatrix(i, 4) = IIf(IsNull(rstrefresh!EX_UNIT), "", rstrefresh!EX_UNIT)
        grdEXPIRYLIST.TextMatrix(i, 5) = IIf(IsNull(rstrefresh!EX_QTY), "", rstrefresh!EX_QTY)
        grdEXPIRYLIST.TextMatrix(i, 6) = IIf(IsNull(rstrefresh!EX_BATCH), "", rstrefresh!EX_BATCH)
        grdEXPIRYLIST.TextMatrix(i, 7) = IIf(IsNull(rstrefresh!EX_DATE), "", rstrefresh!EX_DATE)
        grdEXPIRYLIST.TextMatrix(i, 8) = IIf(IsNull(rstrefresh!EX_MRP), "", Format(rstrefresh!EX_MRP, ".000"))
        grdEXPIRYLIST.TextMatrix(i, 9) = rstrefresh!EX_SLNO
        grdEXPIRYLIST.TextMatrix(i, 10) = IIf(IsNull(rstrefresh!EX_PUR_INV), "", rstrefresh!EX_PUR_INV)
        grdEXPIRYLIST.TextMatrix(i, 11) = IIf(IsNull(rstrefresh!EX_PUR_DATE), "", rstrefresh!EX_PUR_DATE)
        'rstrefresh!EX_SLNO = i
        'rstrefresh.Update
        rstrefresh.MoveNext
    Loop
    rstrefresh.Close
    Set rstrefresh = Nothing
     
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub EXPIRYReport()

    Dim N As Long
    
    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
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
    
    Print #1, Space(3) & AlignLeft(" SL", 2) & Space(1) & _
            AlignLeft("ITEM NAME", 11) & Space(12) & _
            AlignLeft("EXP DATE", 12) & _
            AlignLeft("QTY", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
            
    Print #1, Space(7) & RepeatString("-", 118)
    
    For N = 0 To grdcount.Rows - 1
        Print #1, Space(2) & AlignRight(Str(N + 1), 3) & Space(2) & _
                AlignLeft(grdcount.TextMatrix(N, 1), 20) & Space(5) & _
                AlignLeft(grdcount.TextMatrix(N, 2), 7) & Space(2) & _
                AlignRight(grdcount.TextMatrix(N, 4), 4) & Space(1) & _
                Chr(27) & Chr(72)  '//Bold Ends
    Next N
    
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
    Print #1, Chr(13)
    
    Close #1 '//Closing the file
    
End Sub

Private Function fillcount()
    Dim i, N As Long
    
    grdcount.Rows = 0
    i = 0
    On Error GoTo ErrHand
    For N = 1 To grdEXPIRYLIST.Rows - 1
        If grdEXPIRYLIST.TextMatrix(N, 13) = "Y" Then
            grdcount.Rows = grdcount.Rows + 1
            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(N, 0)
            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(N, 1)
            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(N, 2)
            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(N, 3)
            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(N, 4)
            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(N, 5)
            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(N, 6)
            grdcount.TextMatrix(i, 7) = grdEXPIRYLIST.TextMatrix(N, 7)
            grdcount.TextMatrix(i, 8) = grdEXPIRYLIST.TextMatrix(N, 8)
            grdcount.TextMatrix(i, 9) = grdEXPIRYLIST.TextMatrix(N, 9)
            grdcount.TextMatrix(i, 10) = grdEXPIRYLIST.TextMatrix(N, 10)
            grdcount.TextMatrix(i, 11) = grdEXPIRYLIST.TextMatrix(N, 11)
            grdcount.TextMatrix(i, 12) = grdEXPIRYLIST.TextMatrix(N, 12)
            grdcount.TextMatrix(i, 13) = grdEXPIRYLIST.TextMatrix(N, 13)
            grdcount.TextMatrix(i, 14) = grdEXPIRYLIST.TextMatrix(N, 14)
            i = i + 1
        End If
    Next N
    Exit Function
ErrHand:
    MsgBox Err.Description
    
End Function

Private Function fillList()
    On Error GoTo ErrHand
    grdEXPIRYLIST.ColWidth(0) = 400
    grdEXPIRYLIST.ColWidth(1) = 0
    grdEXPIRYLIST.ColWidth(2) = 2000
    grdEXPIRYLIST.ColWidth(3) = 500
    grdEXPIRYLIST.ColWidth(4) = 500
    grdEXPIRYLIST.ColWidth(5) = 700
    grdEXPIRYLIST.ColWidth(6) = 600
    grdEXPIRYLIST.ColWidth(7) = 650
    grdEXPIRYLIST.ColWidth(8) = 600
    grdEXPIRYLIST.ColWidth(9) = 800
    grdEXPIRYLIST.ColWidth(10) = 1100
    grdEXPIRYLIST.ColWidth(11) = 1200
    
    grdEXPIRYLIST.TextMatrix(0, 0) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 1) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 2) = "MFGR"
    grdEXPIRYLIST.TextMatrix(0, 3) = "SUPPLIER"
    grdEXPIRYLIST.TextMatrix(0, 4) = "PACK"
    grdEXPIRYLIST.TextMatrix(0, 5) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 6) = "BATCH"
    grdEXPIRYLIST.TextMatrix(0, 7) = "EXPIRY"
    grdEXPIRYLIST.TextMatrix(0, 8) = "MRP"
    grdEXPIRYLIST.TextMatrix(0, 9) = "NO"
    grdEXPIRYLIST.TextMatrix(0, 10) = "INV NO"
    grdEXPIRYLIST.TextMatrix(0, 11) = "INV DATE"
    
'    grdEXPIRYLIST.ColWidth(12) = 0
'    grdEXPIRYLIST.ColWidth(13) = 0
'    grdEXPIRYLIST.ColWidth(14) = 0
'    grdEXPIRYLIST.ColWidth(15) = 0
    
    grdEXPIRYLIST.ColAlignment(7) = 3
    grdEXPIRYLIST.ColAlignment(8) = 3
    
    PHYFLAG = True
    TMPFLAG = True
    BATCH_FLAG = True
    ITEM_FLAG = True
    
    txtinvno.Enabled = False
    TXTINVDATE.Enabled = False
    TXTPRODUCT.Enabled = True
    TXTQTY.Enabled = False
    txtBatch.Enabled = False
    M_EDIT = False
    Width = 16110
    Height = 9765
    Left = 0
    Top = 0
    Call FILLEXPIRYGRID
    grdcount.Rows = 0
    TXTSLNO.Text = grdEXPIRYLIST.Rows
    Exit Function
ErrHand:
    MsgBox Err.Description
End Function

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ErrHand
    Select Case KeyCode
        Case vbKeyReturn
                If Not (IsDate(TXTEXPIRY.Text)) Then Exit Sub
                If Len(TXTEXPIRY.Text) < 10 Then Exit Sub
                If DateValue(TXTEXPIRY.Text) > DateValue(Date) Then
                    MsgBox "From Address could not be higher than Today", vbOKOnly, "IMEI REQUEST..."
                    TXTEXPIRY.SetFocus
                    Exit Sub
                End If
                
                
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 2)) & " AND LINE_NO = " & Val(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (rststock.EOF And rststock.BOF) Then
                    rststock!BILL_DATE = Format(TXTEXPIRY.Text, "dd/mm/yyyy")
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = Format(TXTEXPIRY.Text, "dd/mm/yyyy")
                    rststock.Update
                End If
                rststock.Close
                Set rststock = Nothing
                
                grdEXPIRYLIST.Enabled = True
                TXTEXPIRY.Visible = False
                grdEXPIRYLIST.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            grdEXPIRYLIST.SetFocus
    End Select
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Public Function Fillgrid()
    Dim RSTWARTRX As ADODB.Recordset
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand

    Screen.MousePointer = vbHourglass
    
    i = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.Rows = 1
    
    grdEXPIRYLIST.TextMatrix(0, 0) = ""
    grdEXPIRYLIST.TextMatrix(0, 1) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 2) = "REF NO"
    grdEXPIRYLIST.TextMatrix(0, 3) = "DATE"
    grdEXPIRYLIST.TextMatrix(0, 4) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 5) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 6) = "SERIAL NO."
    grdEXPIRYLIST.TextMatrix(0, 7) = "CUSTOMER NAME"
    grdEXPIRYLIST.TextMatrix(0, 8) = "DISTRIBUTOR NAME"
    grdEXPIRYLIST.TextMatrix(0, 9) = "Inv. No"
    grdEXPIRYLIST.TextMatrix(0, 10) = "Inv Date"
    grdEXPIRYLIST.TextMatrix(0, 11) = "ITEM_CODE"
    grdEXPIRYLIST.TextMatrix(0, 12) = "CUSTOMER CODE"
    grdEXPIRYLIST.TextMatrix(0, 13) = "FLAG"
    grdEXPIRYLIST.TextMatrix(0, 14) = "LINE"
    
    grdEXPIRYLIST.ColWidth(0) = 300
    grdEXPIRYLIST.ColWidth(1) = 500
    grdEXPIRYLIST.ColWidth(2) = 1000
    grdEXPIRYLIST.ColWidth(3) = 1500
    grdEXPIRYLIST.ColWidth(4) = 2500
    grdEXPIRYLIST.ColWidth(5) = 800
    grdEXPIRYLIST.ColWidth(6) = 2400
    grdEXPIRYLIST.ColWidth(7) = 2600
    grdEXPIRYLIST.ColWidth(8) = 2600
    grdEXPIRYLIST.ColWidth(9) = 1200
    grdEXPIRYLIST.ColWidth(10) = 1200
    grdEXPIRYLIST.ColWidth(11) = 0
    grdEXPIRYLIST.ColWidth(12) = 0
    grdEXPIRYLIST.ColWidth(13) = 0
    grdEXPIRYLIST.ColWidth(14) = 0
    
    grdEXPIRYLIST.ColAlignment(0) = 9
    grdEXPIRYLIST.ColAlignment(1) = 4
    grdEXPIRYLIST.ColAlignment(2) = 4
    grdEXPIRYLIST.ColAlignment(3) = 4
    grdEXPIRYLIST.ColAlignment(4) = 9
    grdEXPIRYLIST.ColAlignment(5) = 4
    grdEXPIRYLIST.ColAlignment(6) = 1
    grdEXPIRYLIST.ColAlignment(7) = 1
    grdEXPIRYLIST.ColAlignment(8) = 1
    grdEXPIRYLIST.ColAlignment(9) = 4
    grdEXPIRYLIST.ColAlignment(10) = 4
    
    Set RSTWARTRX = New ADODB.Recordset
    With RSTWARTRX
        .Open "SELECT * From WAR_TRXFILE WHERE CHECK_FLAG='N' ORDER BY VCH_NO", db, adOpenForwardOnly
        
        Do Until .EOF
            i = i + 1
            grdEXPIRYLIST.Rows = grdEXPIRYLIST.Rows + 1
            grdEXPIRYLIST.FixedRows = 1
            grdEXPIRYLIST.TextMatrix(i, 0) = ""
            grdEXPIRYLIST.TextMatrix(i, 1) = i
            grdEXPIRYLIST.TextMatrix(i, 2) = !VCH_NO
            grdEXPIRYLIST.TextMatrix(i, 3) = !VCH_DATE
            grdEXPIRYLIST.TextMatrix(i, 4) = IIf(IsNull(!ITEM_NAME), "", !ITEM_NAME)
            grdEXPIRYLIST.TextMatrix(i, 5) = IIf(IsNull(!QTY), "", !QTY)
            grdEXPIRYLIST.TextMatrix(i, 6) = IIf(IsNull(!REF_NO), "", !REF_NO)
            grdEXPIRYLIST.TextMatrix(i, 7) = IIf(IsNull(!ACT_NAME), "", !ACT_NAME)
            grdEXPIRYLIST.TextMatrix(i, 8) = IIf(IsNull(!DIST_NAME), "", !DIST_NAME)
            grdEXPIRYLIST.TextMatrix(i, 9) = IIf(IsNull(!BILL_NO), "", !BILL_NO)
            grdEXPIRYLIST.TextMatrix(i, 10) = IIf(IsNull(!BILL_DATE), "", !BILL_DATE)
            grdEXPIRYLIST.TextMatrix(i, 11) = IIf(IsNull(!ITEM_CODE), "", !ITEM_CODE)
            grdEXPIRYLIST.TextMatrix(i, 12) = IIf(IsNull(!ACT_CODE), "", !ACT_CODE)
            grdEXPIRYLIST.TextMatrix(i, 13) = "N"
            grdEXPIRYLIST.TextMatrix(i, 14) = !line_no
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
            End With

            
            .MoveNext
        Loop
        .Close
        Set RSTWARTRX = Nothing
    End With
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function
