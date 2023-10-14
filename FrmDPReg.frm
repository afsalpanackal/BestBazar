VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMDPReg 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SAVINGS / DEPOSIT REGISTER"
   ClientHeight    =   8415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17955
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8415
   ScaleWidth      =   17955
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
      Height          =   5520
      Left            =   2655
      TabIndex        =   5
      Top             =   1365
      Visible         =   0   'False
      Width           =   9750
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   4200
         Left            =   30
         TabIndex        =   6
         Top             =   1305
         Width           =   9690
         _ExtentX        =   17092
         _ExtentY        =   7408
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
         TabIndex        =   44
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
         TabIndex        =   43
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   7
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00E8EFF7&
      Height          =   8595
      Left            =   -75
      TabIndex        =   0
      Top             =   -165
      Width           =   18045
      Begin VB.CommandButton Command2 
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
         Left            =   13950
         TabIndex        =   45
         Top             =   885
         Width           =   1695
      End
      Begin VB.CommandButton cmdpymnt 
         Caption         =   "Make Receipt"
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
         Left            =   12210
         TabIndex        =   41
         Top             =   885
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print Ledger for All Accounts"
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
         Left            =   10440
         TabIndex        =   40
         Top             =   885
         Width           =   1695
      End
      Begin VB.CommandButton CmdPrint 
         Caption         =   "&PRINT"
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
         Left            =   7380
         TabIndex        =   35
         Top             =   885
         Width           =   1485
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E8EFF7&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   120
         TabIndex        =   19
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
            TabIndex        =   28
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
            TabIndex        =   27
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
            TabIndex        =   25
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
            TabIndex        =   24
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
            TabIndex        =   23
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Receipt Amt"
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
         Left            =   8940
         TabIndex        =   4
         Top             =   885
         Width           =   1440
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
         Left            =   5865
         TabIndex        =   3
         Top             =   885
         Width           =   1440
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00E8EFF7&
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
         Left            =   120
         TabIndex        =   11
         Top             =   75
         Width           =   5670
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
            Left            =   1485
            TabIndex        =   1
            Top             =   225
            Width           =   3735
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1485
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
            TabIndex        =   26
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Party Name"
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
            Index           =   5
            Left            =   255
            TabIndex        =   18
            Top             =   270
            Width           =   1155
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   12
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   13
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Frame1"
         Height          =   6225
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   17895
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
            ItemData        =   "FrmDPReg.frx":0000
            Left            =   0
            List            =   "FrmDPReg.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   34
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
            TabIndex        =   32
            Top             =   0
            Visible         =   0   'False
            Width           =   1350
         End
         Begin MSMask.MaskEdBox TXTEXPIRY 
            Height          =   285
            Left            =   0
            TabIndex        =   31
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
            TabIndex        =   33
            Top             =   0
            Width           =   17880
            _ExtentX        =   31538
            _ExtentY        =   10954
            _Version        =   393216
            Rows            =   1
            Cols            =   20
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   7020
         TabIndex        =   36
         Top             =   270
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   123076609
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTTO 
         Height          =   390
         Left            =   8850
         TabIndex        =   37
         Top             =   285
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         Format          =   123076609
         CurrentDate     =   40498
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
         Left            =   13275
         TabIndex        =   42
         Top             =   525
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
         Left            =   8565
         TabIndex        =   39
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
         Left            =   5850
         TabIndex        =   38
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
         Left            =   13260
         TabIndex        =   29
         Top             =   225
         Width           =   2835
      End
   End
End
Attribute VB_Name = "FRMDPReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CMBYesNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Exit Sub
    On Error GoTo ERRHAND
    Select Case KeyCode
        Case vbKeyReturn
            If CMBYesNo.ListIndex = -1 Then Exit Sub
            If MsgBox("Are you sure you sure...", vbYesNo + vbDefaultButton2, "Payment !!!") = vbNo Then Exit Sub
            Dim rstTRXMAST As ADODB.Recordset
            Set rstTRXMAST = New ADODB.Recordset
            'rstTRXMAST.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'PY' OR TRX_TYPE = 'CR' OR TRX_TYPE = 'PR') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
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
        MsgBox "please Select Customer from the List", vbOKOnly, "Statement"
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
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='FR' OR TRX_TYPE ='FP') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTTRXFILE.EOF
        Select Case RSTTRXFILE!TRX_TYPE
            Case "FR"
                OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case Else
                OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
        End Select
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
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
    RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & DataList2.BoundText & "' and (TRX_TYPE ='FR' OR TRX_TYPE ='FP') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        'RSTTRXFILE!BAL_AMT = Op_Bal
        Select Case RSTTRXFILE!TRX_TYPE
            Case "FR"
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            Case Else
                BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
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
    
    ReportNameVar = Rptpath & "RptDPStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='FR' OR {DBTPYMT.TRX_TYPE} ='FP') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
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

Private Sub CmdPymnt_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMDPRcpt.LBLSUPPLIER.Caption = DataList2.text
    FRMDPRcpt.lblactcode.Caption = DataList2.BoundText
    'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMDPRcpt.Show
End Sub

Private Sub Command1_Click()
    Dim OP_Sale, OP_Rcpt, Op_Bal As Double
    Dim RSTTRXFILE, rstCustomer As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstCustomer = New ADODB.Recordset
    rstCustomer.Open "SELECT * FROM ACTMAST WHERE (Mid(ACT_CODE, 1, 3)='711')And (LENGTH(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstCustomer.EOF
        Op_Bal = 0
        OP_Sale = 0
        OP_Rcpt = 0
        
        OP_Sale = IIf(IsNull(rstCustomer!OPEN_DB), 0, rstCustomer!OPEN_DB)
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='FR' OR TRX_TYPE ='FP') and INV_DATE < '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
        Do Until RSTTRXFILE.EOF
            Select Case RSTTRXFILE!TRX_TYPE
                Case "FR"
                    OP_Sale = OP_Sale + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    OP_Rcpt = OP_Rcpt + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
            End Select
            RSTTRXFILE.MoveNext
        Loop
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        Op_Bal = OP_Sale - OP_Rcpt
            
        rstCustomer!OPEN_CR = Op_Bal
        
        Dim BAL_AMOUNT As Double
        BAL_AMOUNT = 0
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * from DBTPYMT Where ACT_CODE ='" & rstCustomer!ACT_CODE & "' and (TRX_TYPE ='FR' OR TRX_TYPE ='FPR') AND INV_DATE <='" & Format(DTTo.Value, "yyyy/mm/dd") & "' AND INV_DATE >='" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY INV_DATE, TRX_TYPE, INV_NO", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until RSTTRXFILE.EOF
            'RSTTRXFILE!BAL_AMT = Op_Bal
            Select Case RSTTRXFILE!TRX_TYPE
                Case "FR"
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal + IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                Case Else
                    BAL_AMOUNT = BAL_AMOUNT + Op_Bal - IIf(IsNull(RSTTRXFILE!RCPT_AMT), 0, RSTTRXFILE!RCPT_AMT)
                    'RSTTRXFILE!BAL_AMT = Op_Bal
    
            End Select
            RSTTRXFILE!BAL_AMT = BAL_AMOUNT
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
    
    ReportNameVar = Rptpath & "RptDPStatmnt"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='FR' OR {DBTPYMT.TRX_TYPE} ='FP') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTo.Value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.Value, "MM,DD,YYYY") & " #)"
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

Private Sub Command2_Click()
    If DataList2.BoundText = "" Then Exit Sub
    Me.Enabled = False
    FRMDPPYMNT.LBLSUPPLIER.Caption = DataList2.text
    FRMDPPYMNT.lblactcode.Caption = DataList2.BoundText
    'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
    'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
    'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
    'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
    'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
    FRMDPPYMNT.Show
End Sub

Private Sub Form_Activate()
    Call Fillgrid
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "TYPE"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "Rcpt / Paid Date"
    GRDTranx.TextMatrix(0, 3) = "" '"INV NO"
    GRDTranx.TextMatrix(0, 4) = "Receipt Amt"
    GRDTranx.TextMatrix(0, 5) = "Paid Amt"
    GRDTranx.TextMatrix(0, 6) = "REF NO"
    GRDTranx.TextMatrix(0, 7) = "CR NO"
    GRDTranx.TextMatrix(0, 8) = "TYPE"
    GRDTranx.TextMatrix(0, 14) = "Ch. Date."
    GRDTranx.TextMatrix(0, 15) = "Ch. No"
    GRDTranx.TextMatrix(0, 16) = "Bank"
    GRDTranx.TextMatrix(0, 17) = "Mode"
    
    GRDTranx.TextMatrix(0, 19) = "" 'TYPE
    GRDTranx.ColWidth(0) = 1200
    GRDTranx.ColWidth(1) = 800
    GRDTranx.ColWidth(2) = 1900
    GRDTranx.ColWidth(3) = 0
    GRDTranx.ColWidth(4) = 1400
    GRDTranx.ColWidth(5) = 1400
    GRDTranx.ColWidth(6) = 3500
    GRDTranx.ColWidth(7) = 0
    GRDTranx.ColWidth(8) = 0
    GRDTranx.ColWidth(9) = 0
    GRDTranx.ColWidth(10) = 0
    GRDTranx.ColWidth(11) = 0
    GRDTranx.ColWidth(12) = 0
    GRDTranx.ColWidth(13) = 0
    GRDTranx.ColWidth(14) = 1400
    GRDTranx.ColWidth(15) = 1500
    GRDTranx.ColWidth(16) = 1500
    GRDTranx.ColWidth(17) = 1300
    GRDTranx.ColWidth(19) = 0
    'GRDTranx.ColWidth(8) = 0
    
    GRDTranx.ColAlignment(0) = 1
    GRDTranx.ColAlignment(1) = 4
    GRDTranx.ColAlignment(2) = 4
    GRDTranx.ColAlignment(3) = 4
    GRDTranx.ColAlignment(4) = 4
    GRDTranx.ColAlignment(5) = 4
    GRDTranx.ColAlignment(6) = 1
    GRDTranx.ColAlignment(14) = 4
    GRDTranx.ColAlignment(15) = 4
    GRDTranx.ColAlignment(16) = 4
    GRDTranx.ColAlignment(17) = 4
    
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
    
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 300
    Top = 0
    TXTDEALER.text = " "
    TXTDEALER.text = ""
    
    DTFROM.Value = "01/" & Month(Date) & "/" & Year(Date)
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

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Long
    Dim E_TABLE As String
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            
            If GRDTranx.rows = 1 Then Exit Sub
            If GRDTranx.TextMatrix(GRDTranx.Row, 0) = "Payment" Then
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
            'Exit Sub
            If GRDTranx.rows = 1 Then Exit Sub
            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
                If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                    Select Case GRDTranx.Col
                        Case 15
                            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 17)) = "CASH" Then Exit Sub
                            TXTsample.Visible = True
                            TXTsample.Top = GRDTranx.CellTop
                            TXTsample.Left = GRDTranx.CellLeft
                            TXTsample.Width = GRDTranx.CellWidth
                            TXTsample.Height = GRDTranx.CellHeight
                            TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            TXTsample.SetFocus
                        Case 5, 6
                            TXTsample.Visible = True
                            TXTsample.Top = GRDTranx.CellTop
                            TXTsample.Left = GRDTranx.CellLeft
                            TXTsample.Width = GRDTranx.CellWidth
                            TXTsample.Height = GRDTranx.CellHeight
                            TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            TXTsample.SetFocus
                        Case 14
                            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 17)) = "CASH" Then Exit Sub
                            TXTEXPIRY.Visible = True
                            TXTEXPIRY.Top = GRDTranx.CellTop
                            TXTEXPIRY.Left = GRDTranx.CellLeft
                            TXTEXPIRY.Width = GRDTranx.CellWidth
                            TXTEXPIRY.Height = GRDTranx.CellHeight
                            If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                                TXTEXPIRY.text = "  /  /    "
                            Else
                                TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            End If
                            
                            TXTEXPIRY.SetFocus
                        Case 2
                            If Not (UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Or UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT") Then Exit Sub
                            TXTEXPIRY.Visible = True
                            TXTEXPIRY.Top = GRDTranx.CellTop + 120
                            TXTEXPIRY.Left = GRDTranx.CellLeft + 20
                            TXTEXPIRY.Width = GRDTranx.CellWidth
                            TXTEXPIRY.Height = GRDTranx.CellHeight
                            If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                                TXTEXPIRY.text = "  /  /    "
                            Else
                                TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            End If
                            TXTEXPIRY.SetFocus
                        Case 17
                            CMBYesNo.Visible = True
                            CMBYesNo.Top = GRDTranx.CellTop
                            CMBYesNo.Left = GRDTranx.CellLeft
                            CMBYesNo.Width = GRDTranx.CellWidth
                            'CMBYesNo.Height = GRDTranx.CellHeight
                            On Error Resume Next
                            CMBYesNo.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
'                            If Trim(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = "Yes" Then
'                                CMBYesNo.ListIndex = 0
'                            Else
'                                CMBYesNo.ListIndex = 1
'                            End If
                            CMBYesNo.SetFocus
                    End Select
                End If
            End If
            
            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Then
                If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                    Select Case GRDTranx.Col
                        Case 15
                            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 17)) = "CASH" Then Exit Sub
                            TXTsample.Visible = True
                            TXTsample.Top = GRDTranx.CellTop
                            TXTsample.Left = GRDTranx.CellLeft
                            TXTsample.Width = GRDTranx.CellWidth
                            TXTsample.Height = GRDTranx.CellHeight
                            TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            TXTsample.SetFocus
                        Case 4, 6
                            TXTsample.Visible = True
                            TXTsample.Top = GRDTranx.CellTop
                            TXTsample.Left = GRDTranx.CellLeft
                            TXTsample.Width = GRDTranx.CellWidth
                            TXTsample.Height = GRDTranx.CellHeight
                            TXTsample.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            TXTsample.SetFocus
                        Case 14
                            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 17)) = "CASH" Then Exit Sub
                            TXTEXPIRY.Visible = True
                            TXTEXPIRY.Top = GRDTranx.CellTop
                            TXTEXPIRY.Left = GRDTranx.CellLeft
                            TXTEXPIRY.Width = GRDTranx.CellWidth
                            TXTEXPIRY.Height = GRDTranx.CellHeight
                            If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                                TXTEXPIRY.text = "  /  /    "
                            Else
                                TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            End If
                            
                            TXTEXPIRY.SetFocus
                        Case 2
                            If Not (UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Or UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT") Then Exit Sub
                            TXTEXPIRY.Visible = True
                            TXTEXPIRY.Top = GRDTranx.CellTop + 120
                            TXTEXPIRY.Left = GRDTranx.CellLeft + 20
                            TXTEXPIRY.Width = GRDTranx.CellWidth
                            TXTEXPIRY.Height = GRDTranx.CellHeight
                            If Not (IsDate(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col))) Then
                                TXTEXPIRY.text = "  /  /    "
                            Else
                                TXTEXPIRY.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
                            End If
                            TXTEXPIRY.SetFocus
                        Case 17
                            CMBYesNo.Visible = True
                            CMBYesNo.Top = GRDTranx.CellTop
                            CMBYesNo.Left = GRDTranx.CellLeft
                            CMBYesNo.Width = GRDTranx.CellWidth
                            'CMBYesNo.Height = GRDTranx.CellHeight
                            On Error Resume Next
                            CMBYesNo.text = GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)
'                            If Trim(GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col)) = "Yes" Then
'                                CMBYesNo.ListIndex = 0
'                            Else
'                                CMBYesNo.ListIndex = 1
'                            End If
                            CMBYesNo.SetFocus
                    End Select
                End If
            End If
        Case vbKeyF6
            If DataList2.BoundText = "" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Me.Enabled = False
            FRMDPRcpt.LBLSUPPLIER.Caption = DataList2.text
            FRMDPRcpt.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDPRcpt.Show
    Case vbKeyF7
        'Exit Sub
            If DataList2.BoundText = "" Then Exit Sub
            'If GRDTranx.Rows <= 1 Then Exit Sub
            Enabled = False
            FRMDPPYMNT.LBLSUPPLIER.Caption = DataList2.text
            FRMDPPYMNT.lblactcode.Caption = DataList2.BoundText
            'FRMPYMNTSHORT.lblinvdate.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
            'FRMPYMNTSHORT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
            'FRMPYMNTSHORT.lblbillamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
            'FRMPYMNTSHORT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
            'FRMPYMNTSHORT.lblbalamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
            FRMDPPYMNT.Show
    End Select
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
            If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='FR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FR' AND INV_TRX_TYPE = 'FR' AND TRX_TYPE = 'CR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        db.CommitTrans
                        Call Fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
            ElseIf UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.BeginTrans
                        db.Execute "delete From DBTPYMT WHERE TRX_TYPE='FP' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FP' AND INV_TRX_TYPE = 'FP' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "delete FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
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
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='711')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='711')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenForwardOnly
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
    rstTRANX.Open "SELECT * From DBTPYMT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND (TRX_TYPE = 'FR' OR TRX_TYPE = 'FP') ORDER BY INV_DATE DESC", db, adOpenForwardOnly
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.rows = GRDTranx.rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.Row = i
        GRDTranx.Col = 0

        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
'        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
'        Select Case rstTRANX!CHECK_FLAG
'            Case "Y"
'                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!INV_AMT, "0.00")
'            Case "N"
'                GRDTranx.TextMatrix(i, 6) = Format(rstTRANX!INV_AMT, "0.00")
'        End Select
        Select Case rstTRANX!TRX_TYPE
            Case "FR"
                GRDTranx.TextMatrix(i, 4) = IIf(IsNull(rstTRANX!RCPT_AMT), "", Format(rstTRANX!RCPT_AMT, "0.00"))
                GRDTranx.TextMatrix(i, 0) = "Receipt"
                GRDTranx.CellForeColor = vbBlue
            Case "FP"
                GRDTranx.TextMatrix(i, 5) = IIf(IsNull(rstTRANX!RCPT_AMT), "", Format(rstTRANX!RCPT_AMT, "0.00"))
                GRDTranx.TextMatrix(i, 0) = "Payment"
                GRDTranx.CellForeColor = vbRed
        End Select
        GRDTranx.TextMatrix(i, 6) = IIf(IsNull(rstTRANX!REF_NO), "", rstTRANX!REF_NO)
        GRDTranx.TextMatrix(i, 7) = IIf(IsNull(rstTRANX!CR_NO), "", rstTRANX!CR_NO)
        GRDTranx.TextMatrix(i, 8) = IIf(IsNull(rstTRANX!INV_TRX_TYPE), "PI", rstTRANX!INV_TRX_TYPE)
        GRDTranx.TextMatrix(i, 18) = IIf(IsNull(rstTRANX!TRX_YEAR), "", rstTRANX!TRX_YEAR)
        GRDTranx.TextMatrix(i, 19) = IIf(IsNull(rstTRANX!TRX_TYPE), "", rstTRANX!TRX_TYPE)
        
        Select Case rstTRANX!BANK_FLAG
            Case "Y"
                GRDTranx.TextMatrix(i, 9) = IIf(IsNull(rstTRANX!B_TRX_TYPE), "", rstTRANX!B_TRX_TYPE)
                GRDTranx.TextMatrix(i, 10) = IIf(IsNull(rstTRANX!B_TRX_NO), "", rstTRANX!B_TRX_NO)
                GRDTranx.TextMatrix(i, 11) = IIf(IsNull(rstTRANX!B_BILL_TRX_TYPE), "", rstTRANX!B_BILL_TRX_TYPE)
                GRDTranx.TextMatrix(i, 12) = IIf(IsNull(rstTRANX!B_TRX_YEAR), "", rstTRANX!B_TRX_YEAR)
                GRDTranx.TextMatrix(i, 13) = IIf(IsNull(rstTRANX!BANK_CODE), "", rstTRANX!BANK_CODE)
                Dim rstTRXMAST As ADODB.Recordset
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
                If (rstTRANX!TRX_TYPE = "FR" Or rstTRANX!TRX_TYPE = "FP") Then GRDTranx.TextMatrix(i, 17) = "CASH"
        End Select
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
                        MsgBox "Date could not be higher than Today", vbOKOnly, "Receipt Register..."
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    
                    If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From DBTPYMT WHERE TRX_TYPE='FR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock!INV_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                            
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FR' AND INV_TRX_TYPE = 'FR' AND TRX_TYPE = 'CR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
                    ElseIf UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * From DBTPYMT WHERE TRX_TYPE='FP' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (rststock.EOF And rststock.BOF) Then
                            rststock!RCPT_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock!INV_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                            rststock.Update
                            
                        End If
                        rststock.Close
                        Set rststock = Nothing
                        
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FP' AND INV_TRX_TYPE = 'FP' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
                    End If
                        
                    GRDTranx.Enabled = True
                    TXTEXPIRY.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                    
                Case 14
                    If Not (IsDate(TXTEXPIRY.text)) Then
                        TXTEXPIRY.SetFocus
                        Exit Sub
                    End If
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM BANK_TRX WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (rststock.EOF And rststock.BOF) Then
                        rststock!CHQ_DATE = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                        GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(TXTEXPIRY.text, "dd/mm/yyyy")
                        rststock.Update
                    End If
                    rststock.Close
                    Set rststock = Nothing
                        
                    'db.Execute "Update BANK_TRX SET CHQ_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                        
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
                Case 4
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
                    If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "RECEIPT" Then
                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='FR' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FR' AND INV_TRX_TYPE = 'FR' AND TRX_TYPE = 'CR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                    End If
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 5
                    If Val(TXTsample.text) = 0 Then Exit Sub
                    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
                    If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='FP' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FP' AND INV_TRX_TYPE = 'FP' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                    End If
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Format(Val(TXTsample.text), "0.00")
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 15
                    'If RIM(TXTsample.text) = 0 Then Exit Sub
                    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
'                    If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
'                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='FP' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
'                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FP' AND INV_TRX_TYPE = 'FP' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
'                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
'                    End If
                    db.Execute "Update BANK_TRX SET CHQ_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                    'CHQ_NO
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
                Case 6
                    'If RIM(TXTsample.text) = 0 Then Exit Sub
                    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
'                    If UCase(GRDTranx.TextMatrix(GRDTranx.Row, 0)) = "PAYMENT" Then
'                        db.Execute "Update DBTPYMT SET RCPT_AMT = " & Val(TXTsample.text) & " WHERE TRX_TYPE='FP' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
'                        db.Execute "Update CASHATRXFILE SET AMOUNT = " & Val(TXTsample.text) & " WHERE INV_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND INV_TYPE = 'FP' AND INV_TRX_TYPE = 'FP' AND TRX_TYPE = 'DR' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
'                        db.Execute "Update BANK_TRX SET TRX_AMOUNT = " & Val(TXTsample.text) & " WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
'                    End If
                    db.Execute "Update DBTPYMT SET REF_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE= '" & GRDTranx.TextMatrix(GRDTranx.Row, 19) & "' AND CR_NO = " & GRDTranx.TextMatrix(GRDTranx.Row, 7) & " AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 18) & "'"
                    db.Execute "Update BANK_TRX SET REF_NO = '" & Trim(TXTsample.text) & "' WHERE TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 9) & "' AND TRX_NO = " & Val(GRDTranx.TextMatrix(GRDTranx.Row, 10)) & " AND BILL_TRX_TYPE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 11) & "' AND TRX_YEAR = '" & GRDTranx.TextMatrix(GRDTranx.Row, 12) & "' AND BANK_CODE = '" & GRDTranx.TextMatrix(GRDTranx.Row, 13) & "' "
                    
                    
                    GRDTranx.TextMatrix(GRDTranx.Row, GRDTranx.Col) = Trim(TXTsample.text)
                    GRDTranx.Enabled = True
                    TXTsample.Visible = False
                    GRDTranx.SetFocus
                    Exit Sub
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
