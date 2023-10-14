VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMSTKSUMMRY 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK SUMMARY"
   ClientHeight    =   9855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   19125
   ClipControls    =   0   'False
   Icon            =   "FRMSTKSUMRY.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9855
   ScaleWidth      =   19125
   Begin VB.CheckBox chkunbill 
      Caption         =   "Un Billed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   17445
      TabIndex        =   63
      Top             =   8670
      Visible         =   0   'False
      Width           =   1440
   End
   Begin VB.TextBox TxtCode 
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
      Left            =   600
      TabIndex        =   0
      Top             =   15
      Width           =   1830
   End
   Begin VB.TextBox TxtCategory 
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
      Left            =   15435
      TabIndex        =   4
      Top             =   15
      Width           =   3060
   End
   Begin VB.CommandButton CMDRESET 
      Caption         =   "RESET"
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
      Left            =   17865
      TabIndex        =   27
      Top             =   7800
      Width           =   1230
   End
   Begin VB.TextBox tXTNAME1 
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
      Left            =   7905
      TabIndex        =   2
      Top             =   15
      Width           =   2820
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Assign Code"
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
      Left            =   17865
      TabIndex        =   25
      Top             =   7320
      Width           =   1230
   End
   Begin VB.TextBox txtbarcode 
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
      Left            =   11640
      TabIndex        =   3
      Top             =   15
      Width           =   2760
   End
   Begin MSMask.MaskEdBox TXTEXPIRY 
      Height          =   360
      Left            =   0
      TabIndex        =   54
      Top             =   690
      Visible         =   0   'False
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   635
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
   Begin VB.TextBox tXTMEDICINE 
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
      Left            =   3570
      TabIndex        =   1
      Top             =   15
      Width           =   4320
   End
   Begin VB.CommandButton CMDPRINTlABELS 
      Caption         =   "PRINT &LABELS"
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
      Left            =   13935
      TabIndex        =   26
      Top             =   7785
      Width           =   1275
   End
   Begin VB.CheckBox ChkDetails 
      Caption         =   "Detailed"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   17445
      TabIndex        =   40
      Top             =   8340
      Value           =   1  'Checked
      Width           =   1440
   End
   Begin VB.CommandButton CmdPrint 
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
      Height          =   435
      Left            =   15225
      TabIndex        =   23
      Top             =   7335
      Width           =   1305
   End
   Begin VB.CommandButton CmdDisplay 
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
      Height          =   435
      Left            =   13935
      TabIndex        =   22
      Top             =   7335
      Width           =   1275
   End
   Begin VB.Frame Frame 
      Height          =   2190
      Left            =   3750
      TabIndex        =   36
      Top             =   2970
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton CmdCancel 
         Caption         =   "&Cancel"
         Height          =   405
         Left            =   2640
         TabIndex        =   31
         Top             =   1665
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   1335
         TabIndex        =   30
         Top             =   1665
         Width           =   1200
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Commission Type"
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
         Height          =   1470
         Left            =   75
         TabIndex        =   37
         Top             =   150
         Width           =   3780
         Begin VB.TextBox TxtComper 
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
            Left            =   1470
            TabIndex        =   11
            Top             =   765
            Width           =   1650
         End
         Begin VB.Label Label1 
            BackColor       =   &H00000000&
            Caption         =   "Commission"
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
            Height          =   285
            Index           =   24
            Left            =   195
            TabIndex        =   38
            Top             =   765
            Width           =   1260
         End
         Begin MSForms.OptionButton OptAmt 
            Height          =   300
            Left            =   2025
            TabIndex        =   29
            Top             =   315
            Width           =   1140
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2011;529"
            Value           =   "1"
            Caption         =   "Amount"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
         Begin MSForms.OptionButton OptPercent 
            Height          =   300
            Left            =   120
            TabIndex        =   28
            Top             =   315
            Width           =   1365
            VariousPropertyBits=   746588179
            BackColor       =   -2147483633
            ForeColor       =   8388608
            DisplayStyle    =   5
            Size            =   "2408;529"
            Value           =   "0"
            Caption         =   "Percentage"
            FontName        =   "Tahoma"
            FontEffects     =   1073741825
            FontHeight      =   180
            FontCharSet     =   0
            FontPitchAndFamily=   2
            FontWeight      =   700
         End
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
      Left            =   6195
      TabIndex        =   33
      Top             =   1545
      Visible         =   0   'False
      Width           =   1350
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
      Left            =   16560
      TabIndex        =   24
      Top             =   7320
      Width           =   1230
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   6945
      Left            =   45
      TabIndex        =   32
      Top             =   375
      Width           =   19050
      _ExtentX        =   33602
      _ExtentY        =   12250
      _Version        =   393216
      Rows            =   1
      Cols            =   23
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   410
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frmeall 
      BackColor       =   &H00FFC0C0&
      Height          =   2580
      Left            =   60
      TabIndex        =   39
      Top             =   7245
      Width           =   13860
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   465
         Left            =   10500
         TabIndex        =   45
         Top             =   195
         Width           =   3315
         Begin VB.OptionButton OptStock 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Stock Items Only"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   1380
            TabIndex        =   19
            Top             =   150
            Value           =   -1  'True
            Width           =   1890
         End
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00FFC0C0&
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   45
            TabIndex        =   18
            Top             =   165
            Width           =   1335
         End
      End
      Begin VB.TextBox TXTDEALER2 
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   3195
         TabIndex        =   9
         Top             =   405
         Width           =   3015
      End
      Begin VB.CheckBox CHKCATEGORY2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   3180
         TabIndex        =   8
         Top             =   135
         Width           =   1290
      End
      Begin VB.CheckBox CHKCATEGORY 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Supplier"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   60
         TabIndex        =   5
         Top             =   135
         Width           =   1665
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   60
         TabIndex        =   6
         Top             =   405
         Width           =   3030
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   780
         Left            =   60
         TabIndex        =   7
         Top             =   750
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   1376
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataList DataList1 
         Height          =   780
         Left            =   3195
         TabIndex        =   10
         Top             =   750
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   1376
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         Height          =   465
         Left            =   10500
         TabIndex        =   49
         Top             =   570
         Width           =   3315
         Begin VB.OptionButton OptBatch 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Detailed"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   1740
            TabIndex        =   21
            Top             =   135
            Width           =   1425
         End
         Begin VB.OptionButton OptSummary 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Summary"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   45
            TabIndex        =   20
            Top             =   135
            Value           =   -1  'True
            Width           =   1530
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by......"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   6240
         TabIndex        =   53
         Top             =   135
         Width           =   4230
         Begin VB.OptionButton OptName 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Item Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   12
            Top             =   285
            Value           =   -1  'True
            Width           =   1830
         End
         Begin VB.OptionButton OptCode 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Item Code"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   2430
            TabIndex        =   13
            Top             =   300
            Width           =   1620
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sort by......"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1425
         Left            =   6240
         TabIndex        =   46
         Top             =   120
         Width           =   4275
         Begin VB.OptionButton OptLow 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Low Price"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   2085
            TabIndex        =   17
            Top             =   1005
            Width           =   2160
         End
         Begin VB.OptionButton OptHighest 
            BackColor       =   &H00FFC0C0&
            Caption         =   "High Price"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   2085
            TabIndex        =   15
            Top             =   660
            Width           =   2160
         End
         Begin VB.OptionButton OptDead 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Dead moving items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   2100
            TabIndex        =   48
            Top             =   285
            Width           =   2160
         End
         Begin VB.OptionButton Optfast 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Fast moving Items"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   16
            Top             =   1005
            Width           =   2085
         End
         Begin VB.OptionButton OptCategory 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Category"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   14
            Top             =   660
            Width           =   2325
         End
         Begin VB.OptionButton OptSortName 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   315
            Left            =   30
            TabIndex        =   47
            Top             =   285
            Value           =   -1  'True
            Width           =   2070
         End
      End
      Begin VB.Label lblsalval 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   510
         Left            =   11730
         TabIndex        =   62
         Top             =   1065
         Width           =   2070
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Sale Value"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Index           =   9
         Left            =   10350
         TabIndex        =   61
         Top             =   1125
         Width           =   1500
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   8
      Left            =   13905
      TabIndex        =   60
      Top             =   8910
      Width           =   1500
   End
   Begin VB.Label lblnetamt 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   15435
      TabIndex        =   59
      Top             =   8790
      Width           =   1995
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Code"
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
      Index           =   7
      Left            =   15
      TabIndex        =   58
      Top             =   45
      Width           =   525
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
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
      Left            =   14445
      TabIndex        =   57
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   4
      Left            =   10800
      TabIndex        =   56
      Top             =   60
      Width           =   1125
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
      ForeColor       =   &H00FF0000&
      Height          =   300
      Index           =   3
      Left            =   2460
      TabIndex        =   55
      Top             =   60
      Width           =   1125
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F5- Barcode"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   375
      Index           =   2
      Left            =   16620
      TabIndex        =   52
      Top             =   7785
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F4 - Item Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004040&
      Height          =   375
      Index           =   1
      Left            =   15315
      TabIndex        =   51
      Top             =   7995
      Width           =   2460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "F3- Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   0
      Left            =   15315
      TabIndex        =   50
      Top             =   7785
      Width           =   2460
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   44
      Top             =   840
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   43
      Top             =   480
      Width           =   495
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   -570
      TabIndex        =   42
      Top             =   15
      Width           =   495
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   0
      TabIndex        =   41
      Top             =   1080
      Width           =   1620
   End
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   510
      Left            =   15435
      TabIndex        =   35
      Top             =   8235
      Width           =   1995
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   13905
      TabIndex        =   34
      Top             =   8355
      Width           =   1500
   End
End
Attribute VB_Name = "FRMSTKSUMMRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Dim PHY_REC As New ADODB.Recordset
Dim PHY_FLAG As Boolean
Dim M_EDIT As Boolean
Dim Target As Object
Dim LastSave As String 'To Store last Saved Directory
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub CHKCATEGORY2_Click()
    If CHKCATEGORY2.Value = 0 Then
        TXTDEALER2.text = ""
    Else
        TXTDEALER2.SetFocus
    End If
End Sub

Private Sub CHKCATEGORY_Click()
    If chkcategory.Value = 0 Then
        TXTDEALER.text = ""
    Else
        TXTDEALER.SetFocus
    End If
End Sub

Private Sub cmdcancel_Click()
        FRAME.Visible = False
        GRDSTOCK.SetFocus
End Sub

Private Sub CmdCancel_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub CmDDisplay_Click()
    If chkcategory.Value = 1 And DataList2.BoundText = "" Then
        MsgBox "Select Supplier from the List", vbOKOnly, "Stock Register"
        DataList2.SetFocus
        Exit Sub
    End If
    
    If CHKCATEGORY2.Value = 1 And DataList1.BoundText = "" Then
        MsgBox "Select Category from the List", vbOKOnly, "Stock Register"
        DataList1.SetFocus
        Exit Sub
    End If
    
    If OptSummary.Value = True Then
        GRDSTOCK.Cols = 22
        GRDSTOCK.TextMatrix(0, 0) = "SL"
        GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
        GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
        GRDSTOCK.TextMatrix(0, 3) = "QTY"
        GRDSTOCK.TextMatrix(0, 4) = "BARCODE"
        GRDSTOCK.TextMatrix(0, 5) = "Price"
        GRDSTOCK.TextMatrix(0, 6) = "W.Rate"
        GRDSTOCK.TextMatrix(0, 7) = "MRP"
        GRDSTOCK.TextMatrix(0, 8) = "Rcvd Qty"
        GRDSTOCK.TextMatrix(0, 9) = "Sold Qty"
        GRDSTOCK.TextMatrix(0, 10) = "COST"
        GRDSTOCK.TextMatrix(0, 11) = "Total Value"
        GRDSTOCK.TextMatrix(0, 12) = "TRX TYPE"
        GRDSTOCK.TextMatrix(0, 13) = "VCH NO"
        GRDSTOCK.TextMatrix(0, 14) = "LINE NO"
        GRDSTOCK.TextMatrix(0, 15) = "COMISION"
        GRDSTOCK.TextMatrix(0, 16) = "COMI TYPE"
        GRDSTOCK.TextMatrix(0, 17) = "Manufacturer"
        GRDSTOCK.TextMatrix(0, 18) = "Category"
        GRDSTOCK.TextMatrix(0, 19) = "Location"
        
        GRDSTOCK.ColWidth(0) = 700
        'GRDSTOCK.ColWidth(1) = 1500
        GRDSTOCK.ColWidth(2) = 4800
        GRDSTOCK.ColWidth(3) = 800
        GRDSTOCK.ColWidth(4) = 1800
        GRDSTOCK.ColWidth(5) = 1000
        GRDSTOCK.ColWidth(6) = 0
        GRDSTOCK.ColWidth(7) = 1000
        GRDSTOCK.ColWidth(8) = 1000
        GRDSTOCK.ColWidth(9) = 1000
        GRDSTOCK.ColWidth(10) = 1100
        GRDSTOCK.ColWidth(11) = 1400
        GRDSTOCK.ColWidth(12) = 0
        GRDSTOCK.ColWidth(13) = 0
        GRDSTOCK.ColWidth(14) = 0
        GRDSTOCK.ColWidth(15) = 0
        GRDSTOCK.ColWidth(16) = 0
        GRDSTOCK.ColWidth(17) = 2300
        GRDSTOCK.ColWidth(18) = 2300
        GRDSTOCK.ColWidth(19) = 1400
        
        GRDSTOCK.ColAlignment(0) = 1
        GRDSTOCK.ColAlignment(1) = 1
        GRDSTOCK.ColAlignment(2) = 1
        GRDSTOCK.ColAlignment(3) = 4
    '    GRDSTOCK.ColAlignment(4) = 4
    '    GRDSTOCK.ColAlignment(5) = 1
    '    GRDSTOCK.ColAlignment(6) = 4
        GRDSTOCK.ColAlignment(17) = 1
        GRDSTOCK.ColAlignment(18) = 1
    Else
        GRDSTOCK.Cols = 26
        GRDSTOCK.TextMatrix(0, 0) = "SL"
        GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
        GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
        GRDSTOCK.TextMatrix(0, 3) = "QTY"
        GRDSTOCK.TextMatrix(0, 4) = "" '"PACK"
        GRDSTOCK.TextMatrix(0, 5) = "Price"
        GRDSTOCK.TextMatrix(0, 6) = "WS"
        GRDSTOCK.TextMatrix(0, 7) = "VP"
        GRDSTOCK.TextMatrix(0, 8) = "MRP"
        GRDSTOCK.TextMatrix(0, 9) = "EXPIRY"
        GRDSTOCK.TextMatrix(0, 10) = "COST"
        GRDSTOCK.TextMatrix(0, 11) = "Total Value"
        GRDSTOCK.TextMatrix(0, 12) = "TRX TYPE"
        GRDSTOCK.TextMatrix(0, 13) = "VCH NO"
        GRDSTOCK.TextMatrix(0, 14) = "LINE NO"
        GRDSTOCK.TextMatrix(0, 15) = "COMISION"
        GRDSTOCK.TextMatrix(0, 16) = "COMI TYPE"
        GRDSTOCK.TextMatrix(0, 17) = ""
        GRDSTOCK.TextMatrix(0, 18) = ""
        GRDSTOCK.TextMatrix(0, 19) = ""
        GRDSTOCK.TextMatrix(0, 20) = ""
        GRDSTOCK.TextMatrix(0, 21) = "REF"
        GRDSTOCK.TextMatrix(0, 22) = "Barcode"
        
        
        GRDSTOCK.ColWidth(0) = 700
        'GRDSTOCK.ColWidth(1) = 1500
        GRDSTOCK.ColWidth(2) = 4600
        GRDSTOCK.ColWidth(3) = 800
        GRDSTOCK.ColWidth(4) = 0
        GRDSTOCK.ColWidth(5) = 1100
        GRDSTOCK.ColWidth(6) = 1100
        GRDSTOCK.ColWidth(7) = 1100
        GRDSTOCK.ColWidth(8) = 1100
        GRDSTOCK.ColWidth(9) = 1100
        GRDSTOCK.ColWidth(10) = 1100
        GRDSTOCK.ColWidth(11) = 1100
        GRDSTOCK.ColWidth(12) = 0
        GRDSTOCK.ColWidth(13) = 0
        GRDSTOCK.ColWidth(14) = 0
        GRDSTOCK.ColWidth(15) = 0
        GRDSTOCK.ColWidth(16) = 0
        GRDSTOCK.ColWidth(17) = 0
        GRDSTOCK.ColWidth(18) = 0
        GRDSTOCK.ColWidth(19) = 0
        GRDSTOCK.ColWidth(20) = 0
        GRDSTOCK.ColWidth(21) = 1000
        GRDSTOCK.ColWidth(22) = 2500
        GRDSTOCK.ColWidth(23) = 400
        If frmLogin.rs!Level = "0" Then
            GRDSTOCK.TextMatrix(0, 24) = "Net Cost"
            GRDSTOCK.TextMatrix(0, 25) = "Net Amount"
            GRDSTOCK.ColWidth(24) = 1000
            GRDSTOCK.ColWidth(25) = 1000
        Else
            GRDSTOCK.TextMatrix(0, 24) = ""
            GRDSTOCK.TextMatrix(0, 25) = ""
            GRDSTOCK.ColWidth(24) = 0
            GRDSTOCK.ColWidth(25) = 0
        End If
        GRDSTOCK.ColAlignment(0) = 1
        GRDSTOCK.ColAlignment(1) = 1
        GRDSTOCK.ColAlignment(2) = 1
        GRDSTOCK.ColAlignment(3) = 4
    '    GRDSTOCK.ColAlignment(4) = 4
    '    GRDSTOCK.ColAlignment(5) = 1
    '    GRDSTOCK.ColAlignment(6) = 4
        GRDSTOCK.ColAlignment(17) = 1
        GRDSTOCK.ColAlignment(18) = 1
        GRDSTOCK.ColAlignment(20) = 4
        GRDSTOCK.ColAlignment(21) = 1
        GRDSTOCK.ColAlignment(22) = 1
        GRDSTOCK.ColAlignment(24) = 1
        GRDSTOCK.ColAlignment(25) = 1

    End If
    
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim rststock As ADODB.Recordset
    
    If Not IsNumeric(TxtComper.text) Then
        MsgBox " Enter proper value", vbOKOnly, "Commission !!!"
        TxtComper.SetFocus
        Exit Sub
    End If
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (rststock.EOF And rststock.BOF) Then
        If Val(TxtComper.text) = 0 Then
            rststock!COM_FLAG = ""
            rststock!COM_PER = 0
            rststock!COM_AMT = 0
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = "0.00"
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = ""
        Else
            If OptAmt.Value = True Then
                rststock!COM_FLAG = "A"
                rststock!COM_PER = 0
                rststock!COM_AMT = Val(TxtComper.text)
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = Format(Val(TxtComper.text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = "Rs"
            Else
                rststock!COM_FLAG = "P"
                rststock!COM_PER = Val(TxtComper.text)
                rststock!COM_AMT = 0
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 15) = Format(Val(TxtComper.text), "0.00")
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = "%"
            End If
        End If
        rststock.Update
    End If
    rststock.Close
    Set rststock = Nothing
    GRDSTOCK.Enabled = True
    FRAME.Visible = False
End Sub

Private Sub cmdOK_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim i As Integer
    
    If ChkDetails.Value = 0 Then
        ReportNameVar = Rptpath & "RPTSTOCKSMRY"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "( {ITEMMAST.CLOSE_QTY}<> 0 )"
    Else
        ReportNameVar = Rptpath & "RPTDET_STOCK"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "( {RTRXFILE.BAL_QTY}> 0 )"
    End If
    
    
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Address}" Then CRXFormulaField.text = "'LEATHER WORLD' & chr(13) & 'ADOOR'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Stock Report'"
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.Value & "' & ' TO ' &'" & DTTO.Value & "'"
    Next
    frmreport.Caption = "STOCK REPORT"
    Call GENERATEREPORT
End Sub

Private Sub cMDpRINTlABELS_Click()
    Dim n, sl, M As Long
    If GRDSTOCK.rows <= 1 Then Exit Sub
    If GRDSTOCK.Cols = 22 Then Exit Sub
    
    On Error GoTo ERRHAND
    db.Execute "Delete from barprint"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    sl = Val(InputBox("Enter the Serial No. from which to be Print", "Label Printing", 1))
    If sl = 0 Then Exit Sub
    For n = sl To GRDSTOCK.rows - 1
        sl = Val(InputBox("No. of labes for " & Trim(GRDSTOCK.TextMatrix(n, 2)) & " to be printed.", "Label Printing", Val(GRDSTOCK.TextMatrix(n, 3))))
        If sl = 0 Then Exit For
        For M = 1 To sl
            RSTTRXFILE.AddNew
            RSTTRXFILE!BARCODE = "*" & GRDSTOCK.TextMatrix(n, 22) & "*"
            RSTTRXFILE!ITEM_NAME = Trim(GRDSTOCK.TextMatrix(n, 2))
            RSTTRXFILE!item_Price = Val(GRDSTOCK.TextMatrix(n, 5))
            RSTTRXFILE!item_MRP = Val(GRDSTOCK.TextMatrix(n, 7))
            RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).text)
            RSTTRXFILE.Update
        Next M
'            Select Case (MsgBox("Do you want to print Label for " & GRDSTOCK.TextMatrix(N, 2), vbYesNoCancel, "Label Printing!!!"))
'                Case vbYes
'                    'GRDSTOCK.TextMatrix(N, 5)
''                    Picture5.Tag = ""
''                    Picture5.Cls
''                    Picture5.Picture = Nothing
''                    Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
''                    Picture5.CurrentY = 0 'Y2 + 0.25 * Th
''                    Picture5.Print Picture5.Tag & " " & Picture4.Tag
'
'                    Picture5.Cls
'                    Picture5.Picture = Nothing
'                    Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture5.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture5.Print "PRICE: " & Format(GRDSTOCK.TextMatrix(N, 5), "0.00")
'
'                    Picture6.Cls
'                    Picture6.Picture = Nothing
'                    Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture6.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture6.Print "MRP  : " & Format(GRDSTOCK.TextMatrix(N, 7), "0.00")
'
'                    Picture1.Cls
'                    Picture1.Picture = Nothing
'                    Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                    Picture1.CurrentY = 0 'Y2 + 0.25 * Th
'                    Picture1.Print Mid(Trim(GRDSTOCK.TextMatrix(N, 2)), 1, 11) & " MRP: " & Format(GRDSTOCK.TextMatrix(N, 7), "0.00")
'
'                    Dim i As Long
'                    i = Val(InputBox("Enter number of lables to be print", "No. of labels..", GRDSTOCK.TextMatrix(N, 41)))
'                    'i = Val(GRDSTOCK.TextMatrix(N, 41))
'                    If i <= 0 Then Exit Sub
'                    If MDIMAIN.barcode_profile.Caption = 0 Then
'                        If i > 0 Then Call print_3labels(i, Trim(GRDSTOCK.TextMatrix(N, 38)), Trim(GRDSTOCK.TextMatrix(N, 2)), Val(GRDSTOCK.TextMatrix(N, 6)), Val(GRDSTOCK.TextMatrix(N, 18)))
'                        'GRDSTOCK.TextMatrix(Val(TXTSLNO.Text), 6)
'                        '(i As Long, BAR_LABEL As String, itemname As String, itemmrp As Double, itemprice As Double)
'                    Else
'                        If i > 0 Then Call print_labels(i, Trim(GRDSTOCK.TextMatrix(N, 38)), Trim(GRDSTOCK.TextMatrix(N, 2)), Val(GRDSTOCK.TextMatrix(N, 6)), Val(GRDSTOCK.TextMatrix(N, 18)))
'                        'If i > 0 Then Call print_labels(i, Trim(txtBarcode.Text), "")
'                    End If
'                    'Call print_labels(Val(GRDSTOCK.TextMatrix(N, 3)))
'                Case vbCancel
'                    Exit For
'                Case vbNo
'
'            End Select
    Next n
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
          
    Dim i As Integer
    ReportNameVar = Rptpath & "Rptbarprn"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields

    For n = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(n).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(n).Name & " ")
            Report.Database.SetDataSource oRs, 3, n
            Set oRs = Nothing
        End If
    Next n
    
    Set Printer = Printers(barcodeprinter)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    Report.PrintOut (False)
    Set CRXFormulaFields = Nothing
    Set crxApplication = Nothing
    Set Report = Nothing
    Exit Sub
        
Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CMDRESET_Click()
    chkunbill.Visible = False
'    Dim N, sl As Long
'    If GRDSTOCK.Rows <= 1 Then Exit Sub
'    If GRDSTOCK.Cols = 22 Then Exit Sub
'    sl = Val(InputBox("Enter the Serial No. from which to be Print", "Label Printing"))
'    For N = sl To GRDSTOCK.Rows - 1
'        Select Case (MsgBox("Do you want to print Label for " & GRDSTOCK.TextMatrix(N, 2), vbYesNoCancel, "Label Printing!!!"))
'            Case vbYes
'                'GRDSTOCK.TextMatrix(N, 5)
'                 Picture5.Tag = ""
'                Picture4.Tag = ""
'                If Trim(GRDSTOCK.TextMatrix(N, 20)) <> "" Then Picture5.Tag = "Size: " & Trim(GRDSTOCK.TextMatrix(N, 20))
'                If Trim(GRDSTOCK.TextMatrix(N, 21)) <> "" Then Picture4.Tag = "Col: " & Trim(GRDSTOCK.TextMatrix(N, 21))
'                Picture5.Cls
'                Picture5.Picture = Nothing
'                Picture5.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                Picture5.CurrentY = 0 'Y2 + 0.25 * Th
'                Picture5.Print Picture5.Tag & " " & Picture4.Tag
'
'                Picture4.Cls
'                Picture4.Picture = Nothing
'                Picture4.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                Picture4.CurrentY = 0 'Y2 + 0.25 * Th
'                Picture4.Print "PRICE: " & Format(GRDSTOCK.TextMatrix(N, 5), "0.00")
'
'                Picture6.Cls
'                Picture6.Picture = Nothing
'                Picture6.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                Picture6.CurrentY = 0 'Y2 + 0.25 * Th
'                Picture6.Print "MRP  : " & Format(GRDSTOCK.TextMatrix(N, 7), "0.00")
'
'                Picture1.Cls
'                Picture1.Picture = Nothing
'                Picture1.CurrentX = 0 '(Picture1.ScaleWidth - tw) / 2
'                Picture1.CurrentY = 0 'Y2 + 0.25 * Th
'                Picture1.Print Mid(Trim(GRDSTOCK.TextMatrix(N, 2)), 1, 11) & " MRP: " & Format(GRDSTOCK.TextMatrix(N, 7), "0.00")
'
'                Dim i As Long
'                i = Val(InputBox("Enter number of lables to be print", "No. of labels..", GRDSTOCK.TextMatrix(N, 3)))
'                If i <= 0 Then Exit Sub
'                Call print_labels(i, GRDSTOCK.TextMatrix(N, 22), Trim(GRDSTOCK.TextMatrix(N, 21)))
'                'Call print_labels(Val(GRDSTOCK.TextMatrix(N, 3)))
'            Case vbCancel
'                Exit For
'            Case vbNo
'
'        End Select
'    Next N
    
    
    Dim rststock As ADODB.Recordset
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    rststock.Open "SELECT * FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rststock.EOF
        rststock!EDIT_FLAG = ""
        rststock.Update
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    Call Fillgrid
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub Command1_Click()
    If GRDSTOCK.rows <= 1 Then Exit Sub
    If GRDSTOCK.Cols = 22 Then Exit Sub
    
End Sub

Private Sub Form_Load()
    ACT_FLAG = True
    PHY_FLAG = True
    
    db.Execute "Update itemmast set CESS_PER = 0 WHERE ISNULL(CESS_PER)"
    db.Execute "Update itemmast set CESS_AMT = 0 WHERE ISNULL(CESS_AMT)"
    
    'Picture1.FontSize = 5
'    Picture5.FontSize = 8
'    Picture6.FontSize = 8
'    Picture1.FontSize = 8
    Me.Left = 0
    Me.Top = 0
    OptBatch.Value = True
    If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
        lblpvalue.Visible = False
        LBLNETAMT.Visible = False
        lblsalval.Visible = False
        Label1(6).Visible = False
        Label1(8).Visible = False
    End If
    'Call CMDDISPLAY_Click
    'Me.Height = 10000
    'Me.Width = 14595
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
    If PHY_FLAG = False Then PHY_REC.Close
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
'    MDIMAIN.PCTMENU.Enabled = True
'    'MDIMAIN.PCTMENU.Height = 555
'    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDSTOCK_Click()
    If GRDSTOCK.Cols = 22 Then Exit Sub
    TXTsample.Visible = False
    FRAME.Visible = False
    TXTEXPIRY.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub GRDSTOCK_DblClick()
    Dim i As Long
    Dim M As Long
    If GRDSTOCK.rows <= 1 Then Exit Sub
'    If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22)) = "" Then
'
'    End If
    If M_EDIT = True Then
        MsgBox "Changes have been made on MRP Please Refresh the List", vbOKOnly, "Stock Summary"
        Exit Sub
    End If
    If Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)) > 0 Then
        i = Val(InputBox("Enter number of lables to be print", "No. of labels..", GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)))
    Else
        i = Val(InputBox("Enter number of lables to be print", "No. of labels.."))
    End If
    If i <= 0 Then Exit Sub
    
    On Error GoTo ERRHAND
    db.Execute "Delete from barprint"
    Dim RSTTRXFILE As ADODB.Recordset
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From barprint", db, adOpenStatic, adLockOptimistic, adCmdText
    db.BeginTrans
    For M = 1 To i
        RSTTRXFILE.AddNew
        If GRDSTOCK.Cols = 22 Then
            RSTTRXFILE!BARCODE = "*" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4) & "*"
            RSTTRXFILE!ITEM_NAME = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2))
            RSTTRXFILE!item_Price = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5))
            RSTTRXFILE!item_MRP = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
        Else
            RSTTRXFILE!BARCODE = "*" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22) & "*"
            RSTTRXFILE!ITEM_NAME = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 2))
            RSTTRXFILE!item_Price = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5))
            RSTTRXFILE!item_MRP = Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 8))
        End If
        RSTTRXFILE!COMP_NAME = Trim(MDIMAIN.StatusBar.Panels(5).text)
        RSTTRXFILE.Update
    Next M
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    db.CommitTrans
    
    ReportNameVar = Rptpath & "Rptbarprn"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields

    For M = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(M).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(M).Name & " ")
            Report.Database.SetDataSource oRs, 3, M
            Set oRs = Nothing
        End If
    Next M
    
    Set Printer = Printers(barcodeprinter)
    Report.SelectPrinter Printer.DriverName, Printer.DeviceName, Report.PortName
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    Report.PrintOut (False)
    Set CRXFormulaFields = Nothing
    Set crxApplication = Nothing
    Set Report = Nothing
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    
'    If MDIMAIN.StatusBar.Panels(6).Text = "Y" Then
'        If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22)) = "" Then GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22) = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22))
'        If MDIMAIN.barcode_profile.Caption = 0 Then
'            Call print_3labels(i, GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22), "")
'        Else
'            Call print_labels(i, GRDSTOCK.TextMatrix(GRDSTOCK.Row, 22), "")
'        End If
'    End If

    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
            Select Case GRDSTOCK.Col
                'Case 3 '' balQty
                
                Case 3
                    If frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" Then Exit Sub
                    If GRDSTOCK.Cols = 22 Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 350
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 21
                    If GRDSTOCK.Cols = 22 Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 350
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 22
                    If GRDSTOCK.Cols = 22 Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 350
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    If Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)) = "" Then
                        TXTsample.text = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & Int(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)))
                    Else
                        TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    End If
                    TXTsample.SetFocus
                    
                Case 5, 6
                    If frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 350
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                Case 7, 8
                    If GRDSTOCK.Cols = 22 Then Exit Sub
                    If frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4" Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 350
                    TXTsample.Left = GRDSTOCK.CellLeft + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
                    
                Case 9
                    If GRDSTOCK.Cols = 22 Then Exit Sub
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = GRDSTOCK.CellTop + 375
                    TXTEXPIRY.Left = GRDSTOCK.CellLeft + 50
                    TXTEXPIRY.Width = GRDSTOCK.CellWidth '- 25
                    TXTEXPIRY.text = IIf(IsDate(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)), Format(GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col), "MM/YY"), "  /  ")
                    TXTEXPIRY.SetFocus
                Case 15
                    FRAME.Visible = True
                    FRAME.Top = GRDSTOCK.CellTop - 800
                    FRAME.Left = GRDSTOCK.CellLeft - 1500
                    'Frame.Width = GRDSTOCK.CellWidth - 25
                    TxtComper.text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    If GRDSTOCK.TextMatrix(GRDSTOCK.Row, 16) = "Rs" Then
                        OptAmt.Value = True
                    Else
                        OptPercent.Value = True
                    End If
                    TxtComper.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Name...?", "STOCK"))
            For i = 1 To GRDSTOCK.rows - 1
                    If UCase(Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem))) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
        Case 115
            sitem = UCase(InputBox("Item Code...?", "STOCK"))
            For i = 1 To GRDSTOCK.rows - 1
                    If UCase(Mid(GRDSTOCK.TextMatrix(i, 1), 1, Len(sitem))) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
        Case 116
            sitem = UCase(InputBox("Barcode...?", "STOCK"))
            For i = 1 To GRDSTOCK.rows - 1
                    If UCase(Mid(GRDSTOCK.TextMatrix(i, 22), 1, Len(sitem))) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
        Case vbKeyDelete
            If GRDSTOCK.Col <> 23 Then Exit Sub
            Dim rststock As ADODB.Recordset
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = ""
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rststock.EOF And rststock.BOF) Then
                rststock!EDIT_FLAG = "N"
                rststock.Update
            End If
            rststock.Close
            Set rststock = Nothing
            
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
    FRAME.Visible = False
    GRDSTOCK.SetFocus
End Sub

Private Sub Label1_DblClick(index As Integer)
    If chkunbill.Visible = True Then
        chkunbill.Visible = False
    Else
        chkunbill.Visible = True
    End If
End Sub

Private Sub OptAmt_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
             TxtComper.SetFocus
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub OptBatch_Click()
    'CHKCATEGORY.Visible = False
    CHKCATEGORY2.Visible = False
    'TXTDEALER.Visible = False
    TXTDEALER2.Visible = False
    DataList1.Visible = False
    'DataList2.Visible = False
    Frame3.Visible = False
    'Frame2.Visible = False
End Sub

Private Sub OptPercent_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtComper.SetFocus
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub OptSummary_Click()
    chkcategory.Visible = True
    CHKCATEGORY2.Visible = True
    TXTDEALER.Visible = True
    TXTDEALER2.Visible = True
    DataList1.Visible = True
    DataList2.Visible = True
    Frame3.Visible = True
    'Frame2.Visible = True
End Sub

Private Sub txtbarcode_GotFocus()
    TxtBarcode.SelStart = 0
    TxtBarcode.SelLength = Len(TxtBarcode.text)
End Sub

Private Sub txtbarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmDDisplay_Click
    End Select
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmDDisplay_Click
    End Select

End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtName1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmDDisplay_Click
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
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3  ' Bal QTY
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!BAL_QTY = Val(TXTsample.text)
                        rststock!EDIT_FLAG = "Y"
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, 23) = "*"
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    M_STOCK = 0
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenForwardOnly
                    Do Until rststock.EOF
                        M_STOCK = M_STOCK + rststock!BAL_QTY
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
            
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    Do Until rststock.EOF
                        rststock!CLOSE_QTY = M_STOCK
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = TXTsample.text
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 10)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 5  'RT
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_RETAIL = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 6  'WS
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_WS = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 8  'MRP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing

                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 7  'VP
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_VAN = Val(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!P_VAN = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
'                Case 9  'CRTN
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!P_CRTN = Val(TXTsample.Text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
'
''                    Set rststock = New ADODB.Recordset
''                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
''                    If Not (rststock.EOF And rststock.BOF) Then
''                        rststock!P_CRTN = Val(TXTsample.Text)
''                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
''                        rststock.Update
''                    End If
''                    rststock.Close
''                    Set rststock = Nothing
'                    GRDSTOCK.Enabled = True
'                    TXTsample.Visible = False
'                    GRDSTOCK.SetFocus
                    
                Case 10  'COST
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!ITEM_COST = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!ITEM_COST = Val(TXTsample.Text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.Text), "0.000")
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                    
                Case 7  'MRP
                    If Val(TXTsample.text) = 0 Then Exit Sub
                            
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!MRP = Val(TXTsample.text)
                        rststock!P_RETAIL = Val(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(Val(TXTsample.text), "0.000")
                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5) = Format(Val(TXTsample.Text), "0.000")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                    
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!MRP = Val(TXTsample.Text)
'                        rststock!P_RETAIL = Val(TXTsample.Text)
'                        rststock!P_WS = Val(TXTsample.Text)
'                        rststock!P_VAN = Val(TXTsample.Text)
'
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Format(rststock!MRP, "0.000")
'
'                        ''Format(Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) - Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 6)) * 15 / 100, ".000")
'                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) / Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 4)), 2)
'                        ''rststock!P_RETAIL = Round(Val(TXTsample.Text) - Val(TXTsample.Text) * 15 / 100, 2)
'                        ''grdsTOCK.TextMatrix(grdsTOCK.Row, 8) = Format(rststock!P_RETAIL, "0.00")
'                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13) = Format(rststock!P_RETAIL * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
'                        'GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) = Format(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 11)) * Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 3)), "0.00")
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
                    Call TOTALVALUE
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                
                Case 22  'Barcode
                    If Trim(TXTsample.text) = "" Then
                        TXTsample.text = Trim(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1)) & Int(Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 5)))
                    End If
'                    Dim rstTRXMAST As ADODB.Recordset
'                    Set rstTRXMAST = New ADODB.Recordset
'                    rstTRXMAST.Open "Select * From RTRXFILE WHERE BARCODE= '" & Trim(TXTsample.Text) & "' AND ITEM_CODE <> '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' ", db, adOpenStatic, adLockReadOnly
'                    If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
'                        MsgBox "This BARCODE is already being assigned to another Item", vbOKOnly, "Barcode Entry"
'                        TXTsample.SetFocus
'                        rstTRXMAST.Close
'                        Set rstTRXMAST = Nothing
'                        Exit Sub
'                    End If
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
                    
'                    Set rststock = New ADODB.Recordset
'                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                    If Not (rststock.EOF And rststock.BOF) Then
'                        rststock!BARCODE = Trim(TXTsample.text)
'                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
'                        rststock.Update
'                    End If
'                    rststock.Close
'                    Set rststock = Nothing
                    db.Execute "Update RTRXFILE set BARCODE = '" & Trim(TXTsample.text) & "' WHERE ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'"
                    If Trim(TXTsample.text) <> "" Then db.Execute "Update ITEMMAST set BARCODE = '" & Trim(TXTsample.text) & "' where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "'"
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                    
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
                Case 21  'REF
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!REF_NO = Trim(TXTsample.text)
                        GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.text)
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                
                    GRDSTOCK.Enabled = True
                    TXTsample.Visible = False
                    GRDSTOCK.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
        Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Function STOCKADJUST()
'    Dim rststock As ADODB.Recordset
'    Dim RSTITEMMAST As ADODB.Recordset
'
'
'    On Error GoTo eRRHAND
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT CLOSE_QTY from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDPOPUPITEM.Columns(0) & "'", db, adOpenStatic, adLockReadOnly
'    Do Until rststock.EOF
'        M_STOCK = M_STOCK + rststock!CLOSE_QTY
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
'
'
'    Exit Function
'
'eRRHAND:
'    MsgBox Err.Description
End Function

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 5, 6, 7, 8, 9, 10
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
'        Case 5
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
'        Case 7
'             Select Case KeyAscii
'                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
    End Select
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Function TOTALVALUE()
    Dim i As Long
    lblpvalue.Caption = ""
    LBLNETAMT.Caption = ""
    lblsalval.Caption = ""
    For i = 1 To GRDSTOCK.rows - 1
        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 11)), "0.00")
        LBLNETAMT.Caption = Format(Val(LBLNETAMT.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
        lblsalval.Caption = Format(Round(Val(lblsalval.Caption) + Val(GRDSTOCK.TextMatrix(i, 5)) * Val(GRDSTOCK.TextMatrix(i, 3)), 2), "0.00")
    Next i
    lblpvalue.Caption = Format(lblpvalue.Caption, "0.00")
    LBLNETAMT.Caption = Format(LBLNETAMT.Caption, "0.00")
    lblsalval.Caption = Format(lblsalval.Caption, "0.00")
    
End Function

Private Sub TxtComper_GotFocus()
    TxtComper.SelStart = 0
    TxtComper.SelLength = Len(TxtComper.text)
End Sub

Private Sub TxtComper_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdOK_Click
        Case vbKeyEscape
            FRAME.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TxtComper_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TxtComper_LostFocus()
    TxtComper.text = Format(TxtComper.text, "0.00")
End Sub

Private Sub TXTDEALER_Change()
    
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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


Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
    chkcategory.Value = 1
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList2.SetFocus
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

Private Sub DataList2_Click()
        
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text

End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER.text) = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
    chkcategory.Value = 1
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub Fillgrid()
    
    On Error GoTo ERRHAND
    db.Execute "Update rtrxfile set MFGR = '' where isnull(MFGR) "
    db.Execute "Update rtrxfile set CATEGORY = '' where isnull(CATEGORY) "
    db.Execute "Update rtrxfile set BARCODE = '' where isnull(BARCODE) "
    db.Execute "Update rtrxfile set BAL_QTY = 0 where isnull(BAL_QTY) "
    db.Execute "Update rtrxfile set M_USER_ID = '' where isnull(M_USER_ID) "
    
    db.Execute "Update itemmast set ISSUE_QTY = 0 where isnull(ISSUE_QTY) "
    db.Execute "Update itemmast set CATEGORY = '' where isnull(CATEGORY) "
    db.Execute "Update itemmast set MANUFACTURER = '' where isnull(MANUFACTURER) "
    
    Dim rststock As ADODB.Recordset
    Dim rstLOCATION As ADODB.Recordset
    
    Dim i As Long
    Dim P_Value As Double
    Dim S_Value As Double
    
    
    
    i = 0
        
    Screen.MousePointer = vbHourglass
    
    S_Value = 0
    P_Value = 0
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    MDIMAIN.vbalProgressBar1.text = "PLEASE WAIT..."
    On Error Resume Next
    GRDSTOCK.FixedRows = 4
    GRDSTOCK.rows = 1
    lblpvalue.Caption = ""
    LBLNETAMT.Caption = ""
    lblsalval.Caption = ""
    On Error GoTo ERRHAND
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM RTRXFILE WHERE  RTRXFILE.CLOSE_QTY > 0 ORDER BY RTRXFILE.ITEM_NAME", DB, adOpenStatic,adLockReadOnly

    If OptSummary.Value = True Then
        If chkcategory.Value = 1 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            Else
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        ElseIf chkcategory.Value = 1 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            Else
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND MANUFACTURER = '" & DataList2.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 1 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            Else
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY = '" & DataList1.BoundText & "' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        ElseIf chkcategory.Value = 0 And CHKCATEGORY2.Value = 0 Then
            If Optall.Value = True Then
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND ATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND ATEGORY <> 'Service Charge' ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            Else
                If OptSortName.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf optCategory.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.CATEGORY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptDead.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf Optfast.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.ISSUE_QTY DESC", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptLow.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL", db, adOpenStatic, adLockReadOnly
                    End If
                ElseIf OptHighest.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM ITEMMAST WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND CATEGORY <> 'Service Charge' AND  CLOSE_QTY <> 0 ORDER BY ITEMMAST.P_RETAIL DESC", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        End If
    Else
        If chkcategory.Value = 1 Then
            If Trim(TxtBarcode.text) = "" Then
                If Optall.Value = True Then
                    If OptName.Value = True Then
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        End If
                    Else
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        End If
                    End If
                Else
                    If OptName.Value = True Then
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        End If
                    Else
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND M_USER_ID = '" & DataList2.BoundText & "' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        End If
                    End If
                End If
            Else
                If Optall.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.BARCODE Like '" & Me.TxtBarcode.text & "%' AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM RTRXFILE WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND M_USER_ID = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
                    End If
                Else
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.BARCODE Like '" & Me.TxtBarcode.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND RTRXFILE.M_USER_ID = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM RTRXFILE WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND RTRXFILE.BAL_QTY <> 0 AND M_USER_ID = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        Else
            If Trim(TxtBarcode.text) = "" Then
                If Optall.Value = True Then
                    If OptName.Value = True Then
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        End If
                    Else
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        End If
                    End If
                Else
                    If OptName.Value = True Then
                        If chkunbill.Value = 0 Then
                            'rststock.Open "SELECT ITEM_CODE, ITEM_NAME, P_RETAIL, BAL_QTY, P_WS, ITEM_COST  FROM ITEMMAST LEFT JOIN RTRXFILE USING(ITEM_CODE) WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.Text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.Text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.Text & "%' AND RTRXFILE.BAL_QTY <> 0 ", db, adOpenStatic, adLockReadOnly
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 ORDER BY RTRXFILE.ITEM_NAME", db, adOpenStatic, adLockReadOnly
                        End If
                    Else
                        If chkunbill.Value = 0 Then
                            rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.CATEGORY Like '%" & Me.txtcategory.text & "%' AND RTRXFILE.ITEM_CODE Like '" & Me.TxtCode.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND RTRXFILE.ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        Else
                            rststock.Open "SELECT * FROM RTRXFILE WHERE CATEGORY Like '%" & Me.txtcategory.text & "%' AND ITEM_CODE Like '" & Me.TxtCode.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_NAME Like '%" & Me.TxtName1.text & "%' AND RTRXFILE.BAL_QTY <> 0 ORDER BY CONVERT(RTRXFILE.ITEM_CODE, SIGNED INTEGER)", db, adOpenStatic, adLockReadOnly
                        End If
                    End If
                End If
            Else
                If Optall.Value = True Then
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.BARCODE Like '" & Me.TxtBarcode.text & "%' ", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM RTRXFILE WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' ", db, adOpenStatic, adLockReadOnly
                    End If
                Else
                    If chkunbill.Value = 0 Then
                        rststock.Open "SELECT * FROM ITEMMAST LEFT JOIN RTRXFILE ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE (ISNULL(ITEMMAST.UN_BILL) OR ITEMMAST.UN_BILL <> 'Y') AND RTRXFILE.BARCODE Like '" & Me.TxtBarcode.text & "%' AND RTRXFILE.BAL_QTY <> 0", db, adOpenStatic, adLockReadOnly
                    Else
                        rststock.Open "SELECT * FROM RTRXFILE WHERE BARCODE Like '" & Me.TxtBarcode.text & "%' AND RTRXFILE.BAL_QTY <> 0", db, adOpenStatic, adLockReadOnly
                    End If
                End If
            End If
        End If
    End If
    Do Until rststock.EOF
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
        i = i + 1
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 5) = IIf(IsNull(rststock!P_RETAIL), "", Format(rststock!P_RETAIL, "0.000"))
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!P_WS), "", Format(rststock!P_WS, "0.000"))
        GRDSTOCK.TextMatrix(i, 10) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST, "0.000"))
        Select Case rststock!COM_FLAG
            Case "P"
                GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!COM_PER), "", Format(rststock!COM_PER, "0.00"))
                GRDSTOCK.TextMatrix(i, 16) = "%"
            Case "A"
                GRDSTOCK.TextMatrix(i, 15) = IIf(IsNull(rststock!COM_AMT), "", Format(rststock!COM_AMT, "0.00"))
                GRDSTOCK.TextMatrix(i, 16) = "Rs"
        End Select
        If OptSummary.Value = True Then
            GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
            GRDSTOCK.TextMatrix(i, 3) = rststock!CLOSE_QTY
            GRDSTOCK.TextMatrix(i, 4) = rststock!BARCODE
            GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
            GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!RCPT_QTY), "", rststock!RCPT_QTY)
            GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!ISSUE_QTY), "", rststock!ISSUE_QTY)
            GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST * rststock!CLOSE_QTY, "0.00"))
            GRDSTOCK.TextMatrix(i, 17) = IIf(IsNull(rststock!MANUFACTURER), "", rststock!MANUFACTURER)
            GRDSTOCK.TextMatrix(i, 18) = IIf(IsNull(rststock!Category), "", rststock!Category)
            GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rststock!BIN_LOCATION), "", rststock!BIN_LOCATION)
            GRDSTOCK.TextMatrix(i, 20) = Format(Round(IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * rststock!CESS_PER / 100)) + IIf(IsNull(rststock!cess_amt), 0, rststock!cess_amt), 3), "0.000") 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
            GRDSTOCK.TextMatrix(i, 21) = Format(Round(Val(GRDSTOCK.TextMatrix(i, 20)) * Val(GRDSTOCK.TextMatrix(i, 3)), 3), "0.000")
            lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 11)), "0.00")
            LBLNETAMT.Caption = Format(Val(LBLNETAMT.Caption) + Val(GRDSTOCK.TextMatrix(i, 21)), "0.00")
            lblsalval.Caption = Format(Round(Val(lblsalval.Caption) + Val(GRDSTOCK.TextMatrix(i, 5)) * Val(GRDSTOCK.TextMatrix(i, 3)), 2), "0.00")
        Else
            If MDIMAIN.lblcategory.Caption = "Y" Then
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME & IIf(IsNull(rststock!Category), "", " (" & rststock!Category & ")")
            Else
                GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
            End If
            GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_VAN), "", Format(rststock!P_VAN, "0.000"))
            GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(rststock!MRP), "", Format(rststock!MRP, "0.000"))
            GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!EXP_DATE), "", rststock!EXP_DATE)
            
            GRDSTOCK.TextMatrix(i, 3) = rststock!BAL_QTY
            If IsNull(rststock!LINE_DISC) Then
                GRDSTOCK.TextMatrix(i, 4) = 1
            Else
                GRDSTOCK.TextMatrix(i, 4) = rststock!LINE_DISC
            End If
            GRDSTOCK.TextMatrix(i, 11) = IIf(IsNull(rststock!ITEM_COST), "", Format(rststock!ITEM_COST * rststock!BAL_QTY, "0.00"))
            'GRDSTOCK.TextMatrix(i, 20) = IIf(IsNull(rststock!ITEM_SIZE), "", rststock!ITEM_SIZE)
            GRDSTOCK.TextMatrix(i, 21) = IIf(IsNull(rststock!REF_NO), "", rststock!REF_NO)
            GRDSTOCK.TextMatrix(i, 22) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
            Select Case rststock!EDIT_FLAG
                Case "Y"
                    GRDSTOCK.TextMatrix(i, 23) = "*"
                Case Else
                    GRDSTOCK.TextMatrix(i, 23) = ""
            End Select
            GRDSTOCK.TextMatrix(i, 12) = IIf(IsNull(rststock!TRX_TYPE), "", rststock!TRX_TYPE)
            GRDSTOCK.TextMatrix(i, 13) = IIf(IsNull(rststock!VCH_NO), "", rststock!VCH_NO)
            GRDSTOCK.TextMatrix(i, 14) = IIf(IsNull(rststock!LINE_NO), "", rststock!LINE_NO)
            'GRDSTOCK.TextMatrix(i, 24) = Format(Round(IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * rststock!CESS_PER / 100)) + IIf(IsNull(rststock!cess_amt), 0, rststock!cess_amt), 3), "0.000") 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
            GRDSTOCK.TextMatrix(i, 24) = IIf(IsNull(rststock!ITEM_NET_COST_PRICE) Or rststock!ITEM_NET_COST_PRICE < Val(GRDSTOCK.TextMatrix(i, 11)), Val(GRDSTOCK.TextMatrix(i, 11)), Format(rststock!ITEM_NET_COST_PRICE, "0.000"))   'IIf(IsNull(rststock!ITEM_COST), 0, rststock!ITEM_COST + (rststock!ITEM_COST * rststock!SALES_TAX / 100)) + IIf(IsNull(rststock!ITEM_COST), 0, (rststock!ITEM_COST * IIf(IsNull(rststock!CESS_PER), 0, rststock!CESS_PER) / 100)) + IIf(IsNull(rststock!cess_amt), 0, rststock!cess_amt) 'IIf(IsNull(rststock!TRX_TOTAL), "", rststock!TRX_TOTAL / IIf(IsNull(rststock!QTY), "1", rststock!QTY))
            GRDSTOCK.TextMatrix(i, 25) = Format(Round(Val(GRDSTOCK.TextMatrix(i, 24)) * Val(GRDSTOCK.TextMatrix(i, 3)), 3), "0.000")
            Set rstLOCATION = New ADODB.Recordset
            rstLOCATION.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            If Not (rstLOCATION.EOF And rstLOCATION.BOF) Then
                GRDSTOCK.TextMatrix(i, 19) = IIf(IsNull(rstLOCATION!BIN_LOCATION), "", rstLOCATION!BIN_LOCATION)
            End If
            rstLOCATION.Close
            Set rstLOCATION = Nothing
            lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 11)), "0.00")
            LBLNETAMT.Caption = Format(Val(LBLNETAMT.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
            lblsalval.Caption = Format(Round(Val(lblsalval.Caption) + Val(GRDSTOCK.TextMatrix(i, 5)) * Val(GRDSTOCK.TextMatrix(i, 3)), 2), "0.00")
            
        End If
        rststock.MoveNext
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
    Loop
    rststock.Close
    Set rststock = Nothing
    
    M_EDIT = False
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub


Private Sub TXTDEALER2_Change()
    
    On Error GoTo ERRHAND
    If flagchange2.Caption <> "1" Then
        If PHY_FLAG = True Then
            PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY <> 'Service Charge' AND CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            PHY_FLAG = False
        Else
            PHY_REC.Close
            PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY <> 'Service Charge' AND CATEGORY Like '" & TXTDEALER2.text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            PHY_FLAG = False
        End If
        If (PHY_REC.EOF And PHY_REC.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = PHY_REC!Category
        End If
        Set Me.DataList1.RowSource = PHY_REC
        DataList1.ListField = "CATEGORY"
        DataList1.BoundColumn = "CATEGORY"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.text)
    CHKCATEGORY2.Value = 1
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
    End Select

End Sub

Private Sub TXTDEALER2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.text = DataList1.text
    LBLDEALER2.Caption = TXTDEALER2.text

End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.text) = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList1.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
            TXTDEALER2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TXTDEALER2.text = LBLDEALER2.Caption
    DataList1.text = TXTDEALER2.text
    Call DataList1_Click
    CHKCATEGORY2.Value = 1
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmDDisplay_Click
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\"), Asc("["), Asc("]")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    Dim M_DATE As Date
    Dim D As Integer
    Dim M As Integer
    Dim Y As Integer
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Mid(TXTEXPIRY.text, 1, 2)) = 0 Then
                GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = ""
                db.Execute "Update RTRXFILE set EXP_DATE = Null WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'"
                TXTEXPIRY.Visible = False
                GRDSTOCK.Enabled = True
                GRDSTOCK.SetFocus
                Exit Sub
            End If
            If Val(Mid(TXTEXPIRY.text, 1, 2)) > 12 Then Exit Sub
            If Val(Mid(TXTEXPIRY.text, 4, 5)) = 0 Then Exit Sub
            
            M = Val(Mid(TXTEXPIRY.text, 1, 2))
            Y = Val(Right(TXTEXPIRY.text, 2))
            Y = 2000 + Y
            M_DATE = "01" & "/" & M & "/" & Y
            D = LastDayOfMonth(M_DATE)
            M_DATE = D & "/" & M & "/" & Y
            
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 1) & "' AND RTRXFILE.VCH_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 13)) & " AND RTRXFILE.LINE_NO = " & Val(GRDSTOCK.TextMatrix(GRDSTOCK.Row, 14)) & " AND TRX_TYPE = '" & GRDSTOCK.TextMatrix(GRDSTOCK.Row, 12) & "'", db, adOpenStatic, adLockOptimistic
            If Not (rststock.EOF And rststock.BOF) Then
                rststock!EXP_DATE = Format(M_DATE, "dd/mm/yyyy")
                'rststock!VCH_DATE = Format(M_DATE, "dd/mm/yyyy")
                rststock.Update
            End If
            rststock.Close
            Set rststock = Nothing
            
            TXTEXPIRY.Visible = False
            GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = M_DATE
            GRDSTOCK.Enabled = True
            GRDSTOCK.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDSTOCK.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.text)
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CmDDisplay_Click
    End Select
End Sub

