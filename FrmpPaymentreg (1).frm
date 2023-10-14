VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMPaymntreg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PAYMENT REGISTER"
   ClientHeight    =   8430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   18975
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmpPaymentreg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   18975
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
      Height          =   5445
      Left            =   1755
      TabIndex        =   11
      Top             =   1680
      Visible         =   0   'False
      Width           =   9735
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4080
         Left            =   30
         TabIndex        =   12
         Top             =   1335
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   7197
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   11
         Left            =   3705
         TabIndex        =   47
         Top             =   720
         Width           =   1140
      End
      Begin VB.Label lblremarks 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3705
         TabIndex        =   46
         Top             =   975
         Width           =   5985
      End
      Begin VB.Label LBLINVDATE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   135
         TabIndex        =   23
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV DATE"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   405
         TabIndex        =   22
         Top             =   735
         Width           =   885
      End
      Begin VB.Label LBLSUPPLIER 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   990
         TabIndex        =   21
         Top             =   315
         Width           =   4410
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   360
         Width           =   810
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2535
         TabIndex        =   16
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   2685
         TabIndex        =   15
         Top             =   735
         Width           =   810
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV NO"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1665
         TabIndex        =   14
         Top             =   735
         Width           =   675
      End
      Begin VB.Label LBLBILLNO 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   1515
         TabIndex        =   13
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0C0FF&
      Height          =   8595
      Left            =   -75
      TabIndex        =   0
      Top             =   -165
      Width           =   19065
      Begin VB.CommandButton CmdCloseAmt 
         Caption         =   "Print Clo. Amt All Suppliers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   10920
         TabIndex        =   8
         Top             =   885
         Width           =   1290
      End
      Begin VB.CommandButton CmdChqRet 
         Caption         =   "Cheque Return"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   16395
         TabIndex        =   59
         Top             =   885
         Width           =   1260
      End
      Begin VB.CheckBox ChkSmry 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Summary"
         Height          =   210
         Left            =   9900
         TabIndex        =   58
         Top             =   285
         Width           =   1515
      End
      Begin VB.CommandButton CmdPrnRcpt 
         Caption         =   "Print &Payment /Dr /Cr Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   8505
         TabIndex        =   6
         Top             =   885
         Width           =   1095
      End
      Begin VB.CommandButton cdmdrcr 
         Caption         =   "Make Dr && Cr Note Entry"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   13320
         TabIndex        =   10
         Top             =   885
         Width           =   1185
      End
      Begin VB.CommandButton CmdPay 
         Caption         =   "&Mark as Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   10425
         TabIndex        =   54
         Top             =   7770
         Width           =   1170
      End
      Begin VB.CommandButton CmdRemove 
         Caption         =   "Mark as &Not Paid"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   11640
         TabIndex        =   53
         Top             =   7755
         Width           =   1185
      End
      Begin VB.CommandButton cmdpymnt 
         Caption         =   "Make Payment"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   12240
         TabIndex        =   9
         Top             =   885
         Width           =   1050
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Ledger for All Suppliers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   9615
         TabIndex        =   7
         Top             =   885
         Width           =   1290
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "PRINT LEDGER"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6315
         TabIndex        =   4
         Top             =   885
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   120
         TabIndex        =   25
         Top             =   7635
         Width           =   9810
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Op. Balance"
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
            Height          =   300
            Index           =   8
            Left            =   1035
            TabIndex        =   34
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label lblOPBal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   990
            TabIndex        =   33
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLPAIDAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   5325
            TabIndex        =   31
            Top             =   435
            Width           =   1875
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paid Amt"
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
            Height          =   240
            Index           =   3
            Left            =   5340
            TabIndex        =   30
            Top             =   150
            Width           =   1875
         End
         Begin VB.Label LBLINVAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   3135
            TabIndex        =   29
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Invoice Amt"
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
            Height          =   300
            Index           =   6
            Left            =   3180
            TabIndex        =   28
            Top             =   150
            Width           =   1905
         End
         Begin VB.Label LBLBALAMT 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
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
            Height          =   315
            Left            =   7395
            TabIndex        =   27
            Top             =   435
            Width           =   1830
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Bal Amt"
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
            Height          =   315
            Index           =   7
            Left            =   7395
            TabIndex        =   26
            Top             =   150
            Width           =   1815
         End
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7470
         TabIndex        =   5
         Top             =   885
         Width           =   1020
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
         Height          =   525
         Left            =   5205
         TabIndex        =   3
         Top             =   885
         Width           =   1095
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0C0FF&
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
         Height          =   1320
         Left            =   105
         TabIndex        =   17
         Top             =   75
         Width           =   5070
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
            Left            =   1260
            TabIndex        =   1
            Top             =   225
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1260
            TabIndex        =   2
            Top             =   585
            Width           =   3735
            _ExtentX        =   6588
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
         Begin VB.PictureBox rptPRINT 
            Height          =   480
            Left            =   9990
            ScaleHeight     =   420
            ScaleWidth      =   1140
            TabIndex        =   32
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
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
            ForeColor       =   &H000000FF&
            Height          =   315
            Index           =   5
            Left            =   75
            TabIndex        =   24
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   18
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   19
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   6225
         Left            =   120
         TabIndex        =   36
         Top             =   1440
         Width           =   18930
         Begin VB.Frame Frmereceipt 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
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
            Height          =   4065
            Left            =   8430
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   4245
            Begin VB.TextBox TXTsample2 
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
               Left            =   1890
               TabIndex        =   56
               Top             =   1050
               Visible         =   0   'False
               Width           =   1350
            End
            Begin MSFlexGridLib.MSFlexGrid grdreceipts 
               Height          =   3825
               Left            =   45
               TabIndex        =   57
               Top             =   195
               Width           =   4155
               _ExtentX        =   7329
               _ExtentY        =   6747
               _Version        =   393216
               Rows            =   1
               Cols            =   6
               FixedRows       =   0
               FixedCols       =   0
               RowHeightMin    =   400
               BackColorFixed  =   0
               ForeColorFixed  =   65535
               BackColorBkg    =   12632256
               AllowBigSelection=   0   'False
               FocusRect       =   2
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
         End
         Begin MSFlexGridLib.MSFlexGrid grdcount 
            Height          =   5190
            Left            =   0
            TabIndex        =   48
            Top             =   0
            Visible         =   0   'False
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   9155
            _Version        =   393216
            Rows            =   1
            Cols            =   26
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
         Begin VB.PictureBox picChecked 
            Height          =   285
            Left            =   0
            Picture         =   "FrmpPaymentreg.frx":030A
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   50
            Top             =   600
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.PictureBox picUnchecked 
            Height          =   285
            Left            =   120
            Picture         =   "FrmpPaymentreg.frx":064C
            ScaleHeight     =   225
            ScaleWidth      =   225
            TabIndex        =   49
            Top             =   0
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.ComboBox CMBYesNo 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            ItemData        =   "FrmpPaymentreg.frx":098E
            Left            =   0
            List            =   "FrmpPaymentreg.frx":0998
            Style           =   2  'Dropdown List
            TabIndex        =   40
            Top             =   0
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox TXTsample 
            Appearance      =   0  'Flat
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
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Top             =   -15
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   285
            Left            =   0
            TabIndex        =   37
            Top             =   720
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
         Begin MSFlexGridLib.MSFlexGrid GRDTranx 
            Height          =   6210
            Left            =   15
            TabIndex        =   39
            Top             =   0
            Width           =   18900
            _ExtentX        =   33338
            _ExtentY        =   10954
            _Version        =   393216
            Rows            =   1
            Cols            =   28
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   450
            BackColorFixed  =   0
            ForeColorFixed  =   65535
            BackColorBkg    =   12632256
            FocusRect       =   2
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
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   6375
         TabIndex        =   41
         Top             =   270
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   104464385
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   8205
         TabIndex        =   42
         Top             =   285
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   104464385
         CurrentDate     =   40498
      End
      Begin VB.Label LBLSelected 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   450
         Left            =   14520
         TabIndex        =   52
         Top             =   930
         Width           =   1770
      End
      Begin VB.Label LBLTOTAL 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   315
         Index           =   12
         Left            =   14460
         TabIndex        =   51
         Top             =   705
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F7 to make Dr/Cr Notes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   240
         Index           =   0
         Left            =   12690
         TabIndex        =   45
         Top             =   450
         Width           =   3030
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
         Index           =   9
         Left            =   7920
         TabIndex        =   44
         Top             =   345
         Width           =   285
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "Period From"
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
         Index           =   10
         Left            =   5205
         TabIndex        =   43
         Top             =   345
         Width           =   1140
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 to make payments"
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
         Index           =   8
         Left            =   12675
         TabIndex        =   35
         Top             =   195
         Width           =   2835
      End
   End
End
Attribute VB_Name = "FRMPaymntreg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub cdmdrcr_Click()
    If DataList2.BoundText = "" Then Exit Sub
    'If GRDTranx.Rows <= 1 Then Exit Sub
    Enabled = False
    FRMDRCR2.LBLSUPPLIER.Caption = DataList2.text
    FRMDRCR2.lblactcode.Caption = DataList2.BoundText
    'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMDRCR2.Show
End Sub

Private Sub CMBYesNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Exit Sub
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If CMBYesNo.ListIndex = -1 Then Exit Sub
            If MsgBox("Are you sure you sure...", vbYesNo + vbDefaultButton2, "Payment !!!") = vbNo Then Exit Sub
            Dim rstTRXMAST As ADODB.Recordset
            Set rstTRXMAST = New ADODB.Recordset
            'rstTRXMAST.Open "SELECT * From CRDTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'PY' OR TRX_TYPE = 'CR' OR TRX_TYPE = 'PR') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
            db.BeginTrans
            rstTRXMAST.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 10) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                If CMBYesNo.ListIndex = 0 Then
                    rstTRXMAST!BANK_FLAG = "N"
                Else
                    rstTRXMAST!BANK_FLAG = "Y"
                End If
                rstTRXMAST.Update
                GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = CMBYesNo.text
                
                'GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rsttrxmast!CHQ_DATE), "", rsttrxmast!CHQ_DATE)
                'GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rsttrxmast!CHQ_NO), "", rsttrxmast!CHQ_NO)
                'GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rsttrxmast!BANK_NAME), "", rsttrxmast!BANK_NAME)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            db.CommitTrans
            CMBYesNo.Visible = False
            GRDTranx.SetFocus
        Case vbKeyEscape
            CMBYesNo.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub CmdChqRet_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMCHQRET.LBLSUPPLIER.Caption = DataList2.text
    FRMCHQRET.lblactcode.Caption = DataList2.BoundText
    'FRMCHQRET.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMCHQRET.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMCHQRET.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMCHQRET.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMCHQRET.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMCHQRET.Show
    FRMCHQRET.SetFocus
End Sub

Private Sub CmdCloseAmt_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, rstCustomer As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        OP_Sale = IIf(IsNull(rstCustomer!OPEN_DB), 0, rstCustomer!OPEN_DB)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMOUNT) from CRDTPYMT WHERE ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE <> 'CR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE <> 'PY' AND TRX_TYPE <> 'PR' AND TRX_TYPE <> 'WP' AND TRX_TYPE <> 'DB' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Op_Bal = OP_Sale - OP_Rcpt
            
        rstCustomer!OPEN_CR = Op_Bal
        
        rstCustomer.MoveNext
    Loop
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    Screen.MousePointer = vbNormal
    Sleep (300)
    
    ReportNameVar = Rptpath & "RptSupStmntAll"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "((Mid({ACTMAST.ACT_CODE}, 1, 3)='311')And (LENGTH({ACTMAST.ACT_CODE})>3))"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM CRDTPYMT ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ACTMAST ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then
            CRXFormulaField.text = "'STATEMENT FOR THE PERIOD FROM ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        End If
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmDDisplay_Click()
    Call Fillgrid
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, RstCustmast As ADODB.Recordset
    Dim i As Long
    
    If DataList2.BoundText = "" Then
        MsgBox "please Select Supplier from the List", vbOKOnly, "Statement"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    Op_Bal = 0
    OP_Sale = 0
    OP_Rcpt = 0
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Sale = IIf(IsNull(RSTTRXFILE!OPEN_DB), 0, RSTTRXFILE!OPEN_DB)
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select SUM(RCPT_AMOUNT) from CRDTPYMT WHERE ACT_CODE ='" & DataList2.BoundText & "' and TRX_TYPE <> 'CR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
                    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "select SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE ='" & DataList2.BoundText & "' and TRX_TYPE <> 'PY' AND TRX_TYPE <> 'PR' AND TRX_TYPE <> 'WP' AND TRX_TYPE <> 'DB' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='RC' OR TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='CR') and INV_DATE < '" & Format(DTFROM.value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'    Do Until RSTTRXFILE.EOF
'        If RSTTRXFILE!TRX_TYPE <> "CR" Then
'            OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
'        End If
'        If Not (RSTTRXFILE!TRX_TYPE = "PY" Or RSTTRXFILE!TRX_TYPE = "PR" Or RSTTRXFILE!TRX_TYPE = "DB") Then
'            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
'        End If
'
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
    Op_Bal = OP_Sale - OP_Rcpt
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * FROM ACTMAST WHERE ACT_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        RSTTRXFILE!OPEN_CR = Op_Bal
        RSTTRXFILE.Update
    End If
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Dim BAL_AMOUNT As Double
    BAL_AMOUNT = 0
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='RC' OR  TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='WP' OR TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='CR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        Select Case RSTTRXFILE!TRX_TYPE
            Case "PY", "PR", "DB", "WP"
                'BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
            Case "CR"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT)
                'RSTTRXFILE!BAL_AMT = Op_Bal

        End Select
        RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        Op_Bal = 0
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    
    Screen.MousePointer = vbNormal
    Sleep (300)
    
    ReportNameVar = Rptpath & "RptSupStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='RC' OR {DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='WP' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM CRDTPYMT ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ACTMAST ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then
            CRXFormulaField.text = "'STATEMENT OF ' & '" & UCase(DataList2.text) & "' & CHR(13) &' FOR THE PERIOD FROM ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        End If
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdPrnRcpt_Click()
    Dim CompName, CompAddress1, CompAddress2, CompAddress3, CompTin, CompCST As String
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        CompName = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        CompAddress1 = IIf(IsNull(RSTCOMPANY!Address), "", RSTCOMPANY!Address)
        CompAddress2 = IIf(IsNull(RSTCOMPANY!HO_NAME), "", RSTCOMPANY!HO_NAME)
        If Trim(CompAddress2) = "" Then
            CompAddress2 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        Else
            CompAddress3 = "Ph: " & IIf(IsNull(RSTCOMPANY!TEL_NO), "", RSTCOMPANY!TEL_NO) & IIf((IsNull(RSTCOMPANY!FAX_NO)) Or RSTCOMPANY!FAX_NO = "", "", ", " & RSTCOMPANY!FAX_NO) & _
                        IIf((IsNull(RSTCOMPANY!EMAIL_ADD)) Or RSTCOMPANY!EMAIL_ADD = "", "", "Email: " & RSTCOMPANY!FAX_NO)
        End If
        CompTin = IIf(IsNull(RSTCOMPANY!CST) Or RSTCOMPANY!CST = "", "", "GSTIN No. " & RSTCOMPANY!CST)
        CompCST = IIf(IsNull(RSTCOMPANY!DL_NO) Or RSTCOMPANY!DL_NO = "", "", "CST No. " & RSTCOMPANY!DL_NO)
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Screen.MousePointer = vbHourglass
    Sleep (300)
    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Then
        ReportNameVar = Rptpath & "RptPymnt"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({CRDTPYMT.TRX_TYPE} ='PY' AND {CRDTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {CRDTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
        ReportNameVar = Rptpath & "RptDNP"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({CRDTPYMT.TRX_TYPE} ='DB' AND {CRDTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {CRDTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
        ReportNameVar = Rptpath & "RptCNP"
        Set Report = crxApplication.OpenReport(ReportNameVar, 1)
        Report.RecordSelectionFormula = "({CRDTPYMT.TRX_TYPE} ='CB' AND {CRDTPYMT.CR_NO} = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND {CRDTPYMT.ACT_CODE}='" & DataList2.BoundText & "')"
    Else
        Screen.MousePointer = vbNormal
        Exit Sub
    End If
    'Report.RecordSelectionFormula = "({CRDTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({CRDTPYMT.TRX_TYPE} ='CB' OR {CRDTPYMT.TRX_TYPE} ='DB' OR {CRDTPYMT.TRX_TYPE} ='RT' OR {CRDTPYMT.TRX_TYPE} ='DR' OR {CRDTPYMT.TRX_TYPE} ='SR')) "
    
    
    Dim i As Integer
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
    Set CRXFormulaFields = Report.FormulaFields
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.text = "'" & CompCST & "'"
        If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.text = "'Statement of ' & '" & UCase(DataList2.text) & "' & ' for the Period ' &'" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub CmdPymnt_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMPYMNTSHORT.LBLSUPPLIER.Caption = DataList2.text
    FRMPYMNTSHORT.lblactcode.Caption = DataList2.BoundText
    'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMPYMNTSHORT.Show
End Sub

Private Sub Command1_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, rstCustomer As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        OP_Sale = IIf(IsNull(rstCustomer!OPEN_DB), 0, rstCustomer!OPEN_DB)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(RCPT_AMOUNT) from CRDTPYMT WHERE ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE <> 'CR' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Rcpt = IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
                        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "select SUM(INV_AMT) from CRDTPYMT WHERE ACT_CODE ='" & rstCustomer!ACT_CODE & "' and TRX_TYPE <> 'PY' AND TRX_TYPE <> 'PR' AND TRX_TYPE <> 'WP' AND TRX_TYPE <> 'DB' and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
            OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE.Fields(0)), 0, RSTTRXFILE.Fields(0))
        End If
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Op_Bal = OP_Sale - OP_Rcpt
            
        rstCustomer!OPEN_CR = Op_Bal
        
        Dim BAL_AMOUNT As Double
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from CRDTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='RC' OR TRX_TYPE ='PY' OR TRX_TYPE ='PR' OR TRX_TYPE ='WP' OR TRX_TYPE ='CB' OR TRX_TYPE ='DB' OR TRX_TYPE ='CR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "PY", "PR", "DB", "WP"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
                Case "CR"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
        
            
'            'RSTTRXFILE!BAL_AMT = Op_Bal
'            Select Case RSTTRXFILE!TRX_TYPE
'                Case "CR"
'                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!INV_AMT), 0, RSTTRXFILE!INV_AMT) '- IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT))
'                Case Else
'                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), 0, RSTTRXFILE!RCPT_AMOUNT)
'                    'RSTTRXFILE!BAL_AMT = Op_Bal
'
'            End Select
'            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
            Op_Bal = 0
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        rstCustomer.MoveNext
    Loop
    rstCustomer.Close
    Set rstCustomer = Nothing
    
    Screen.MousePointer = vbNormal
    Sleep (300)
    
    If ChkSmry.Value = 1 Then
        ReportNameVar = Rptpath & "RptSupStmntSmry"
    Else
        ReportNameVar = Rptpath & "RptSupStatmnt"
    End If
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='RC' OR {DBTPYMT.TRX_TYPE} ='PY' OR {DBTPYMT.TRX_TYPE} ='PR' OR {DBTPYMT.TRX_TYPE} ='WP' OR {DBTPYMT.TRX_TYPE} ='CB' OR {DBTPYMT.TRX_TYPE} ='DB' OR {DBTPYMT.TRX_TYPE} ='CR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
    Next i
    
    If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM CRDTPYMT ")
        Report.Database.SetDataSource oRs, 3, 1
        Set oRs = Nothing
        
        Set oRs = New ADODB.Recordset
        Set oRs = db.Execute("SELECT * FROM ACTMAST ")
        Report.Database.SetDataSource oRs, 3, 2
        Set oRs = Nothing
    End If
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        If CRXFormulaField.Name = "{@PERIOD}" Then
            CRXFormulaField.text = "'STATEMENT FOR THE PERIOD FROM ' & '" & DTFROM.Value & "' & ' TO ' &'" & DTTo.Value & "'"
        End If
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.text = "'" & MDIMAIN.StatusBar.Panels(5).text & "'"
    Next
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "INV / PAID DATE"
    GRDTranx.TextMatrix(0, 3) = "COMP REF"
    GRDTranx.TextMatrix(0, 4) = "INV AMT"
    GRDTranx.TextMatrix(0, 5) = "PAID AMT"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "CR NO"
    GRDTranx.TextMatrix(0, 8) = "TYPE"
    GRDTranx.TextMatrix(0, 14) = "Ch. Date."
    GRDTranx.TextMatrix(0, 15) = "Ch. No"
    GRDTranx.TextMatrix(0, 16) = "Bank"
    GRDTranx.TextMatrix(0, 17) = "Mode"
    GRDTranx.TextMatrix(0, 20) = "Inv No."
    GRDTranx.TextMatrix(0, 21) = "Paid Amt"
    GRDTranx.TextMatrix(0, 22) = "Balance Amt"
    GRDTranx.TextMatrix(0, 23) = "Status"
    GRDTranx.TextMatrix(0, 24) = "Days"
    GRDTranx.TextMatrix(0, 27) = "Remarks"
    GRDTranx.ColWidth(0) = 1000
    GRDTranx.ColWidth(1) = 700
    GRDTranx.ColWidth(2) = 1500
    GRDTranx.ColWidth(3) = 1000
    GRDTranx.ColWidth(4) = 1200
    GRDTranx.ColWidth(5) = 1200
    GRDTranx.ColWidth(6) = 1400
    GRDTranx.ColWidth(7) = 0
    GRDTranx.ColWidth(8) = 0
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 1200
    GRDTranx.ColWidth(15) = 1300
    GRDTranx.ColWidth(16) = 1300
    GRDTranx.ColWidth(17) = 1000
    GRDTranx.ColWidth(18) = 0 '300
    GRDTranx.ColWidth(19) = 0 '300
    GRDTranx.ColWidth(20) = 1100
    GRDTranx.ColWidth(21) = 1100
    GRDTranx.ColWidth(22) = 1100
    GRDTranx.ColWidth(23) = 600
    GRDTranx.ColWidth(24) = 600
    GRDTranx.ColWidth(25) = 0
    GRDTranx.ColWidth(26) = 300
    GRDTranx.ColWidth(27) = 1200
    'GRDTranx.ColWidth(8) = 0
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 4
    GRDTranx.ColAlignment(14) = 4
    GRDTranx.ColAlignment(15) = 4
    GRDTranx.ColAlignment(16) = 4
    GRDTranx.ColAlignment(17) = 4
    GRDTranx.ColAlignment(20) = 4
    GRDTranx.ColAlignment(23) = 4
    GRDTranx.ColAlignment(24) = 4
    GRDTranx.ColAlignment(27) = 1
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "Description"
    GRDBILL.TextMatrix(0, 2) = "Rate"
    GRDBILL.TextMatrix(0, 3) = "Disc %"
    GRDBILL.TextMatrix(0, 4) = "Tax %"
    GRDBILL.TextMatrix(0, 5) = "Qty"
    GRDBILL.TextMatrix(0, 6) = "Amount"
    
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 2800
    GRDBILL.ColWidth(2) = 800
    GRDBILL.ColWidth(3) = 800
    GRDBILL.ColWidth(4) = 800
    GRDBILL.ColWidth(5) = 900
    GRDBILL.ColWidth(6) = 1100
    
    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(2) = 6
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    GRDBILL.ColAlignment(6) = 6
    
    grdreceipts.TextMatrix(0, 0) = "SL"
    grdreceipts.TextMatrix(0, 1) = "Date"
    grdreceipts.TextMatrix(0, 2) = "Amount"
    grdreceipts.TextMatrix(0, 3) = ""
    grdreceipts.TextMatrix(0, 4) = ""
    grdreceipts.TextMatrix(0, 5) = ""
    
    grdreceipts.ColWidth(0) = 500
    grdreceipts.ColWidth(1) = 1500
    grdreceipts.ColWidth(2) = 1800
    grdreceipts.ColWidth(3) = 0
    grdreceipts.ColWidth(4) = 0
    grdreceipts.ColWidth(5) = 0
    
    grdreceipts.ColAlignment(0) = 3
    grdreceipts.ColAlignment(2) = 3
    grdreceipts.ColAlignment(3) = 3
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 0
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    
    If Month(Date) - 2 > 4 Then
        DTFROM.Value = "01/" & Format(Month(Date) - 2, "00") & "/" & Year(Date)
    Else
        If Year(Date) > Year(MDIMAIN.DTFROM.Value) Then
            DTFROM.Value = "01/12/" & Year(MDIMAIN.DTFROM.Value)
        Else
            DTFROM.Value = "01/04/" & Year(MDIMAIN.DTFROM.Value)
        End If
    End If
    DTTo.Value = Format(Date, "DD/MM/YYYY")
    
    'MDIMAIN.MNUPYMNT.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close

    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.MNUPYMNT.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        GRDTranx.SetFocus
    End If
End Sub

Private Sub grdreceipts_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            Frmereceipt.Visible = False
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub grdreceipts_LostFocus()
    Frmereceipt.Visible = False
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim E_TABLE As String
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            
            On Error GoTo ERRHAND
            If GRDTranx.rows = 1 Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "Purchase" Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Purchase" And GRDTranx.Col = 21 Then
                grdreceipts.rows = 1
                i = 0
                
'                        Set rstTRANX2 = New ADODB.Recordset
'                        rstTRANX2.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND INV_NO  = " & rstTRANX!INV_NO & " AND INV_TRX_TYPE = '" & rstTRANX!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                        If Not (rstTRANX2.EOF And rstTRANX2.BOF) Then
'                            rstTRANX!RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0))
'                            rstTRANX.Update
'                            'db.Execute "Update DBTPYMT set RCVD_AMOUNT = IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)) where ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND TRX_TYPE = 'DR' AND INV_TRX_TYPE  = '" & rstTRANX!TRX_TYPE & "' AND INV_NO = '" & rstTRANX!VCH_NO & "' AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "'"
'                            'lblsaleret.Caption = Format(IIf(IsNull(rstTRANX2.Fields(0)), 0, rstTRANX2.Fields(0)), "0.00")
'                        End If
'                        rstTRANX2.Close
'                        Set rstTRANX2 = Nothing

                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * From trnxrcpt WHERE ACT_CODE = '" & DataList2.BoundText & "' AND INV_TRX_YEAR = '" & Val(GRDTranx.TextMatrix(GRDTranx.Row, 18)) & "' AND INV_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 3)) & " AND INV_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' ", db, adOpenForwardOnly
                Do Until RSTTRXFILE.EOF
                    i = i + 1
                    grdreceipts.rows = grdreceipts.rows + 1
                    grdreceipts.FixedRows = 1
                    grdreceipts.TextMatrix(i, 0) = i
                    grdreceipts.TextMatrix(i, 1) = Format(RSTTRXFILE!RCPT_DATE, "dd/mm/yyyy")
                    grdreceipts.TextMatrix(i, 2) = IIf(IsNull(RSTTRXFILE!RCPT_AMOUNT), "0.00", Format(RSTTRXFILE!RCPT_AMOUNT, "0.00"))
                    grdreceipts.TextMatrix(i, 3) = RSTTRXFILE!RCPT_NO
                    grdreceipts.TextMatrix(i, 4) = RSTTRXFILE!TRX_TYPE
                    grdreceipts.TextMatrix(i, 5) = RSTTRXFILE!INV_NO
                    RSTTRXFILE.MoveNext
                Loop
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
    
                Frmereceipt.Visible = True
                grdreceipts.SetFocus
                
                Exit Sub
            Else
                LBLSUPPLIER.Caption = " " & DataList2.text
                LBLINVDATE.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 2)
                LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
                LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 4)
                'LBLPAID.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 5)
                'LBLBAL.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)
    
                GRDBILL.rows = 1
                i = 0
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * From RTRXFILE WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenForwardOnly
                Do Until RSTTRXFILE.EOF
                    i = i + 1
                    GRDBILL.rows = GRDBILL.rows + 1
                    GRDBILL.FixedRows = 1
                    GRDBILL.TextMatrix(i, 0) = i
                    GRDBILL.TextMatrix(i, 1) = RSTTRXFILE!ITEM_NAME
                    GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!SALES_PRICE, "0.00")
                    GRDBILL.TextMatrix(i, 3) = Val(RSTTRXFILE!LINE_DISC)
                    GRDBILL.TextMatrix(i, 4) = Val(RSTTRXFILE!SALES_TAX)
                    GRDBILL.TextMatrix(i, 5) = RSTTRXFILE!QTY
                    GRDBILL.TextMatrix(i, 6) = Format(RSTTRXFILE!TRX_TOTAL, "0.00")
                    RSTTRXFILE.MoveNext
                Loop
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
                
                Set RSTTRXFILE = New ADODB.Recordset
                RSTTRXFILE.Open "Select * From TRANSMAST WHERE VCH_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 8) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenForwardOnly
                If Not (RSTTRXFILE.EOF Or RSTTRXFILE.BOF) Then
                    lblremarks.Caption = IIf(IsNull(RSTTRXFILE!REMARKS), "", RSTTRXFILE!REMARKS)
                End If
                RSTTRXFILE.Close
                Set RSTTRXFILE = Nothing
            End If
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
        
        Case 113
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            Select Case GRDTranx.Col
                 Case 2
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note") Then Exit Sub
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = GRDTranx.CellTop '+ 120
                    TXTEXPIRY.Left = GRDTranx.CellLeft '+ 20
                    TXTEXPIRY.Width = GRDTranx.CellWidth
                    TXTEXPIRY.Height = GRDTranx.CellHeight
                    If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                        TXTEXPIRY.text = "  /  /    "
                    Else
                        TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                    End If
                    TXTEXPIRY.SetFocus
                Case 5
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note") Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop '+ 120
                    TXTsample.Left = GRDTranx.CellLeft '+ 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
                Case 4
                    If Not (GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Or GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note") Then Exit Sub
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop '+ 120
                    TXTsample.Left = GRDTranx.CellLeft '+ 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
                Case 6
                    TXTsample.Visible = True
                    TXTsample.Top = GRDTranx.CellTop '+ 120
                    TXTsample.Left = GRDTranx.CellLeft '+ 20
                    TXTsample.Width = GRDTranx.CellWidth
                    TXTsample.Height = GRDTranx.CellHeight
                    TXTsample.text = Val(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))
                    TXTsample.SetFocus
            End Select
        Case vbKeyF6
            If DataList2.BoundText = "" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Me.Enabled = False
            FRMPYMNTSHORT.LBLSUPPLIER.Caption = DataList2.text
            FRMPYMNTSHORT.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMPYMNTSHORT.Show
    Case vbKeyF7
        'Exit Sub
            If DataList2.BoundText = "" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMDRCR2.LBLSUPPLIER.Caption = DataList2.text
            FRMDRCR2.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDRCR2.Show
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub GRDTranx_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRHAND
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
            If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From CRDTPYMT WHERE TRX_TYPE='PY' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PY' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        db.Execute "delete From trnxrcpt WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.CommitTrans
                        Call Fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From CRDTPYMT WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' "
                        db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'DN' AND INV_TRX_TYPE = 'DN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        db.CommitTrans
                        Call Fillgrid
                    End If
                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From CRDTPYMT WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN'"
                        db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'CN' AND INV_TRX_TYPE = 'CN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        db.CommitTrans
                        Call Fillgrid
                    End If
                End If
            ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From CRDTPYMT WHERE TRX_TYPE='RC' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        'db.Execute "delete From trnxrcpt WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.CommitTrans
                        Call Fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
            End If
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number <> -2147168237 Then
        MsgBox err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub GRDTranx_Scroll()
    Frmereceipt.Visible = False
    FRMEBILL.Visible = False
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            
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

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='311')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ACT_NAME
        End If
        Set DataList2.RowSource = ACT_REC
        DataList2.ListField = "ACT_NAME"
        DataList2.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    GRDTranx.rows = 1
    Call Fillgrid
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Supplier From List", vbOKOnly, "EzBiz"
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim rstTRXMAST As ADODB.Recordset
    Dim i As Long
    
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ERRHAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From CRDTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'RC' OR TRX_TYPE = 'CB' OR TRX_TYPE = 'DB' OR TRX_TYPE = 'PY' OR TRX_TYPE = 'CR' OR TRX_TYPE = 'PR' OR TRX_TYPE = 'WP') ORDER BY INV_DATE, CR_NO", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
        Select Case rstTRANX!check_flag
            Case "Y"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!INV_AMT, "0.00")
            Case "N"
                GRDTranx.TextMatrix(i, 5) = 0 '""Format(rstTRANX!INV_AMT, "0.00")
        End Select
        Select Case rstTRANX!TRX_TYPE
            Case "CR"
                GRDTranx.TextMatrix(i, 20) = IIf(IsNull(rstTRANX!PINV), "", rstTRANX!PINV)
                GRDTranx.TextMatrix(i, 24) = DateDiff("d", rstTRANX!INV_DATE, Date)
                GRDTranx.TextMatrix(i, 0) = "Purchase"
                GRDTranx.CellForeColor = vbRed
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "select SUM(RCPT_AMOUNT) from trnxrcpt WHERE TRX_TYPE = 'PY' AND ACT_CODE = '" & rstTRANX!ACT_CODE & "' AND INV_NO  = " & rstTRANX!INV_NO & " AND INV_TRX_TYPE = '" & rstTRANX!INV_TRX_TYPE & "' AND INV_TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                    GRDTranx.TextMatrix(i, 21) = IIf(IsNull(rstTRXMAST.Fields(0)), 0, rstTRXMAST.Fields(0))
                End If
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                
                GRDTranx.TextMatrix(i, 22) = Val(GRDTranx.TextMatrix(i, 4)) - Val(GRDTranx.TextMatrix(i, 21))
                Select Case rstTRANX!PAID_FLAG
                    Case "Y"
                        GRDTranx.TextMatrix(i, 22) = "0"
                        GRDTranx.TextMatrix(i, 23) = "PAID"
                        GRDTranx.TextMatrix(i, 24) = ""
                    Case Else
                        GRDTranx.TextMatrix(i, 23) = "PEND"
                End Select
            Case "CB"
                GRDTranx.TextMatrix(i, 0) = "Credit Note"
                GRDTranx.CellForeColor = vbRed
            Case "PY"
                GRDTranx.TextMatrix(i, 0) = "Payment"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "PR"
                GRDTranx.TextMatrix(i, 0) = "Purchase Return"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "WP"
                GRDTranx.TextMatrix(i, 0) = "Purchase Return(W)"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "DB"
                GRDTranx.TextMatrix(i, 0) = "Debit Note"
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
                GRDTranx.CellForeColor = vbBlue
            Case "RC"
                GRDTranx.TextMatrix(i, 0) = "Cheque Return"
                'GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMOUNT, "0.00")
                GRDTranx.CellForeColor = vbRed
        End Select
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!INV_TRX_TYPE), "PI", rstTRANX!INV_TRX_TYPE)
        GRDTranx.TextMatrix(i, 18) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        GRDTranx.TextMatrix(i, 19) = IIf(IsNull(rstTRANX!TRX_TYPE), "", rstTRANX!TRX_TYPE)
        GRDTranx.TextMatrix(i, 27) = IIf(IsNull(rstTRANX!REMARKS), "", rstTRANX!REMARKS)
        Select Case rstTRANX!BANK_FLAG
            Case "Y"
                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_TYPE), "", rstTRANX!B_TRX_TYPE)
                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
                GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!B_BILL_TRX_TYPE), "", rstTRANX!B_BILL_TRX_TYPE)
                GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!B_TRX_YEAR), "", rstTRANX!B_TRX_YEAR)
                GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
                Set rstTRXMAST = New ADODB.Recordset
                rstTRXMAST.Open "SELECT * From BANK_TRX WHERE BANK_CODE = '" & GRDTranx.TextMatrix(i, 13) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(i, 12) & "' AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(i, 11) & "' AND TRX_NO = " & GRDTranx.TextMatrix(i, 10) & " AND TRX_TYPE = '" & GRDTranx.TextMatrix(i, 9) & "'  ORDER BY TRX_DATE", db, adOpenForwardOnly
                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                    GRDTranx.TextMatrix(i, 14) = IIf(IsNull(rstTRXMAST!CHQ_DATE), "", rstTRXMAST!CHQ_DATE)
                    GRDTranx.TextMatrix(i, 15) = IIf(IsNull(rstTRXMAST!CHQ_NO), "", rstTRXMAST!CHQ_NO)
                    GRDTranx.TextMatrix(i, 16) = IIf(IsNull(rstTRXMAST!BANK_NAME), "", rstTRXMAST!BANK_NAME)
                End If
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                GRDTranx.TextMatrix(i, 17) = "BANK"
            Case Else
                GRDTranx.TextMatrix(i, 9) = ""
                GRDTranx.TextMatrix(i, 10) = ""
                GRDTranx.TextMatrix(i, 11) = ""
                GRDTranx.TextMatrix(i, 12) = ""
                GRDTranx.TextMatrix(i, 13) = ""
                If rstTRANX!TRX_TYPE = "PY" Then GRDTranx.TextMatrix(i, 17) = "CASH"
        End Select
        
        
        
        With GRDTranx
            If .TextMatrix(.Row, 8) = "PI" Or .TextMatrix(.Row, 8) = "PW" Or .TextMatrix(.Row, 8) = "LP" Then
                .Row = i: .Col = 26: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
        
        If GRDTranx.TextMatrix(i, 0) = "Purchase" And Val(GRDTranx.TextMatrix(i, 22)) <= 0 Then
            GRDTranx.TextMatrix(i, 23) = "PAID"
            GRDTranx.TextMatrix(i, 24) = ""
        End If
        
        If GRDTranx.TextMatrix(i, 23) = "PAID" Then
            GRDTranx.Row = i
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(i, 23) = "PEND" Then
            GRDTranx.Row = i
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbRed
        End If
        
        GRDTranx.Row = i
        GRDTranx.Col = 0
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + Val(GRDTranx.TextMatrix(i, 4)), "0.00")
        LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from ACTMAST  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLBALAMT.Caption = Format(Val(lblOPBal.Caption) + (Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption)), "0.00")
    
    If GRDTranx.rows > 13 Then GRDTranx.TopRow = GRDTranx.rows - 1
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
ERRHAND:
    Screen.MousePointer = vbDefault
    MsgBox err.Description
End Function

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.text)
End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 2
                    If Not (IsDate(TXTEXPIRY.text)) Then
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    If DateValue(TXTEXPIRY.text) > DateValue(Date) Then
                        MsgBox "Date could not be higher than Today", vbOKOnly, "Payment Register..."
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From CRDTPYMT WHERE TRX_TYPE='PY' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock!INV_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                            
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From TRNXRCPT WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PY' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CRDTPYMT WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'DN' AND INV_TRX_TYPE = 'DN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CRDTPYMT WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'CN' AND INV_TRX_TYPE = 'CN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!VCH_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            Set rststock = New ADODB.Recordset
                            rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                            If Not (rststock.EOF And rststock.BOF) Then
                                rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                                rststock.Update
                            End If
                            rststock.Close
                            Set rststock = Nothing
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From CRDTPYMT WHERE TRX_TYPE='RC' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock!INV_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'RC' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!TRX_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                    End If
                        
                    GRDTranx.Enabled = True
                    TXTEXPIRY.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                
            End Select
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.Visible = False
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
   
    
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDTranx.Col
                Case 5
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Then
                        db.Execute "Update CRDTPYMT SET RCPT_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='PY' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "Update TRNXRCPT SET RCPT_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'PY' AND INV_TRX_TYPE = 'PY' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Debit Note" Then
                    
                        db.Execute "Update CRDTPYMT SET RCPT_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='DB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'DN' "
                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'DN' AND INV_TRX_TYPE = 'DN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'CR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'CB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    End If
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 4
                    If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Cheque Return" Then
                        db.Execute "Update CRDTPYMT SET INV_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='RC' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        'db.Execute "Update TRNXRCPT SET RCPT_AMOUNT = " & Val(TXTsample.Text) & " WHERE TRX_TYPE='PY' AND CR_TRX_TYPE= 'CR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " "
                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    ElseIf GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Credit Note" Then
                        db.Execute "Update CRDTPYMT SET INV_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='CB' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TRX_TYPE = 'CN'"
                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'CN' AND INV_TRX_TYPE = 'CN' AND TRX_TYPE = 'CT' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        If Not GRDTranx.TextMatrix(GRDTranx.Row, 10) = "" Then
                            db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = 'DR' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = 'DB' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "'"
                        End If
                        
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                        Call Fillgrid
                    End If
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                
                Case 6  ' Ref No
                    'If Trim(TXTsample.Text) = "" Then Exit Sub
                    db.Execute "Update crdtpymt set REF_NO = '" & Trim(TXTsample.text) & "' where CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & "  and TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' "

                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
            End Select
            
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDTranx.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDTranx.Col
        Case 15, 6
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 5
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

Private Sub CMBYesNo_LostFocus()
    CMBYesNo.Visible = False
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub


Private Function fillcount()
    Dim i, n As Long
    
    grdcount.rows = 0
    i = 0
    LBLSelected.Caption = ""
    On Error GoTo ERRHAND
    For n = 1 To GRDTranx.rows - 1
        If GRDTranx.TextMatrix(n, 25) = "Y" Then
            grdcount.rows = grdcount.rows + 1
            grdcount.TextMatrix(i, 0) = GRDTranx.TextMatrix(n, 0)
            grdcount.TextMatrix(i, 1) = GRDTranx.TextMatrix(n, 1)
            grdcount.TextMatrix(i, 2) = GRDTranx.TextMatrix(n, 2)
            grdcount.TextMatrix(i, 3) = GRDTranx.TextMatrix(n, 3)
            grdcount.TextMatrix(i, 4) = GRDTranx.TextMatrix(n, 4)
            grdcount.TextMatrix(i, 5) = GRDTranx.TextMatrix(n, 5)
            grdcount.TextMatrix(i, 6) = GRDTranx.TextMatrix(n, 6)
            grdcount.TextMatrix(i, 7) = GRDTranx.TextMatrix(n, 7)
            grdcount.TextMatrix(i, 8) = GRDTranx.TextMatrix(n, 8)
            grdcount.TextMatrix(i, 9) = GRDTranx.TextMatrix(n, 9)
            grdcount.TextMatrix(i, 10) = GRDTranx.TextMatrix(n, 10)
            grdcount.TextMatrix(i, 11) = GRDTranx.TextMatrix(n, 11)
            grdcount.TextMatrix(i, 12) = GRDTranx.TextMatrix(n, 12)
            grdcount.TextMatrix(i, 13) = GRDTranx.TextMatrix(n, 13)
            grdcount.TextMatrix(i, 14) = GRDTranx.TextMatrix(n, 14)
            grdcount.TextMatrix(i, 15) = GRDTranx.TextMatrix(n, 15)
            grdcount.TextMatrix(i, 16) = GRDTranx.TextMatrix(n, 16)
            grdcount.TextMatrix(i, 17) = GRDTranx.TextMatrix(n, 17)
            grdcount.TextMatrix(i, 18) = GRDTranx.TextMatrix(n, 18)
            grdcount.TextMatrix(i, 19) = GRDTranx.TextMatrix(n, 19)
            grdcount.TextMatrix(i, 20) = GRDTranx.TextMatrix(n, 20)
            grdcount.TextMatrix(i, 21) = GRDTranx.TextMatrix(n, 21)
            grdcount.TextMatrix(i, 22) = GRDTranx.TextMatrix(n, 22)
            grdcount.TextMatrix(i, 23) = GRDTranx.TextMatrix(n, 23)
            grdcount.TextMatrix(i, 24) = GRDTranx.TextMatrix(n, 24)
            grdcount.TextMatrix(i, 25) = n
            
            
            LBLSelected.Caption = Val(LBLSelected.Caption) + Val(GRDTranx.TextMatrix(n, 22))
            i = i + 1
        End If
    Next n
    
    LBLSelected.Caption = Format(LBLSelected.Caption, "0.00")
    Exit Function
ERRHAND:
    MsgBox err.Description
    
End Function

Private Sub GRDTranx_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If GRDTranx.rows = 1 Then Exit Sub
    If GRDTranx.Col <> 26 Then Exit Sub
    With GRDTranx
        If .TextMatrix(.Row, 0) = "Purchase" Then
            oldx = .Col
            oldy = .Row
            .Row = oldy: .Col = 26: .CellPictureAlignment = 4
                'If GRDTranx.Col = 0 Then
                    If GRDTranx.CellPicture = picChecked Then
                        Set GRDTranx.CellPicture = picUnchecked
                        '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                        'strTextCheck = .Text
                        ' When you de-select a CheckBox, we need to strip out the #
                        'strChecked = strChecked & strTextCheck & ","
                        ' Don't forget to strip off the trailing , before passing the string
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 25) = "Y"
                        Call fillcount
                    Else
                        Set GRDTranx.CellPicture = picChecked
                        '.Col = .Col + 2
                        'strTextCheck = .Text
                        'strChecked = Replace(strChecked, strTextCheck & ",", "")
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 25) = "N"
                        Call fillcount
                    End If
                'End If
            .Col = oldx
            .Row = oldy
        End If
    End With
End Sub

Private Sub CmdPay_Click()
    If MsgBox("Are you sure you want to make the selected invoices as Paid", vbYesNo + vbDefaultButton2, "Payment Entry") = vbNo Then Exit Sub
    
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ERRHAND
    For i = 0 To grdcount.rows - 1
        db.Execute "UPDATE CRDTPYMT SET PAID_FLAG = 'Y' WHERE TRX_TYPE='CR' AND CR_NO = " & grdcount.TextMatrix(i, 7) & " AND INV_TRX_TYPE = '" & grdcount.TextMatrix(i, 8) & "' AND TRX_YEAR= '" & grdcount.TextMatrix(i, 18) & "'"
        GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PAID"
        
        If GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PAID" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 25)
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PEND" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 25)
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbRed
        End If
        
    Next i
    
    For i = 1 To GRDTranx.rows - 1
        GRDTranx.TextMatrix(i, 25) = "N"
        With GRDTranx
            If .TextMatrix(.Row, 8) = "PI" Or .TextMatrix(.Row, 8) = "PW" Or .TextMatrix(.Row, 8) = "LP" Then
                .Row = i: .Col = 26: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
    Next i
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Cmdremove_Click()
    If MsgBox("Are you sure you want to make the selected invoices as Not Paid", vbYesNo + vbDefaultButton2, "Payment Entry") = vbNo Then Exit Sub
    Dim i As Long
    Dim RSTTRXFILE As ADODB.Recordset
    On Error GoTo ERRHAND
    For i = 0 To grdcount.rows - 1
        db.Execute "UPDATE CRDTPYMT SET PAID_FLAG = 'N' WHERE TRX_TYPE='CR' AND CR_NO = " & grdcount.TextMatrix(i, 7) & " AND INV_TRX_TYPE = '" & grdcount.TextMatrix(i, 8) & "' AND TRX_YEAR= '" & grdcount.TextMatrix(i, 18) & "'"
        GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PEND"
        
        If GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PAID" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 25)
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbBlue
        ElseIf GRDTranx.TextMatrix(grdcount.TextMatrix(i, 25), 23) = "PEND" Then
            GRDTranx.Row = grdcount.TextMatrix(i, 25)
            GRDTranx.Col = 23
            GRDTranx.CellForeColor = vbRed
        End If
        
    Next i
    
    For i = 1 To GRDTranx.rows - 1
        GRDTranx.TextMatrix(i, 25) = "N"
        With GRDTranx
            If .TextMatrix(.Row, 8) = "PI" Or .TextMatrix(.Row, 8) = "PW" Or .TextMatrix(.Row, 8) = "LP" Then
                .Row = i: .Col = 26: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
    Next i
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

