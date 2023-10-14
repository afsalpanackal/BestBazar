VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRcpts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "RECEIPT REGISTER"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   11070
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   120
      Picture         =   "FrmRptReg.frx":0000
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   0
      Picture         =   "FrmRptReg.frx":0342
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   41
      Top             =   600
      Visible         =   0   'False
      Width           =   285
   End
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
      Height          =   4830
      Left            =   1140
      TabIndex        =   6
      Top             =   1755
      Visible         =   0   'False
      Width           =   7005
      Begin MSFlexGridLib.MSFlexGrid GRDBILL 
         Height          =   3480
         Left            =   45
         TabIndex        =   7
         Top             =   1305
         Width           =   6915
         _ExtentX        =   12197
         _ExtentY        =   6138
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LBLPAID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   3615
         TabIndex        =   23
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "PAID AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   9
         Left            =   3720
         TabIndex        =   22
         Top             =   735
         Width           =   930
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "BAL AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   8
         Left            =   4935
         TabIndex        =   21
         Top             =   735
         Width           =   870
      End
      Begin VB.Label LBLBAL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   4815
         TabIndex        =   20
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLINVDATE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   45
         TabIndex        =   18
         Top             =   975
         Width           =   1365
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV DATE"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   4
         Left            =   315
         TabIndex        =   17
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
         Left            =   1080
         TabIndex        =   16
         Top             =   315
         Width           =   4125
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "SUPPLIER"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   900
      End
      Begin VB.Label LBLBILLAMT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2445
         TabIndex        =   11
         Top             =   975
         Width           =   1155
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV AMT"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   1
         Left            =   2595
         TabIndex        =   10
         Top             =   735
         Width           =   810
      End
      Begin VB.Label LBLTOTAL 
         BackStyle       =   0  'Transparent
         Caption         =   "INV NO"
         ForeColor       =   &H000080FF&
         Height          =   255
         Index           =   0
         Left            =   1575
         TabIndex        =   9
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
         Left            =   1425
         TabIndex        =   8
         Top             =   975
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEMAIN 
      BackColor       =   &H00C0FFC0&
      Height          =   9540
      Left            =   -75
      TabIndex        =   0
      Top             =   -150
      Width           =   11160
      Begin VB.TextBox TXTREFNO 
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
         ForeColor       =   &H00FF00FF&
         Height          =   345
         Left            =   6900
         MaxLength       =   15
         TabIndex        =   39
         Top             =   1005
         Width           =   2625
      End
      Begin VB.CommandButton CmdPay 
         Caption         =   "&Make Payments"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   8340
         TabIndex        =   38
         Top             =   375
         Width           =   1185
      End
      Begin MSFlexGridLib.MSFlexGrid grdcount 
         Height          =   5190
         Left            =   180
         TabIndex        =   35
         Top             =   2820
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   9155
         _Version        =   393216
         Rows            =   1
         Cols            =   14
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
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "&EXIT"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4905
         TabIndex        =   5
         Top             =   900
         Width           =   1350
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
         Height          =   480
         Left            =   4905
         TabIndex        =   4
         Top             =   345
         Width           =   1350
      End
      Begin VB.Frame Frmeperiod 
         BackColor       =   &H00C0FFC0&
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
         TabIndex        =   12
         Top             =   90
         Width           =   4785
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
            Left            =   1050
            TabIndex        =   1
            Top             =   225
            Width           =   3690
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   645
            Left            =   1050
            TabIndex        =   2
            Top             =   585
            Width           =   3690
            _ExtentX        =   6509
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
            TabIndex        =   24
            Top             =   -45
            Width           =   1200
         End
         Begin VB.Label LBLTOTAL 
            BackStyle       =   0  'Transparent
            Caption         =   "Customer"
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
            TabIndex        =   19
            Top             =   300
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   8010
            TabIndex        =   13
            Top             =   630
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            BackColor       =   &H00FF80FF&
            Height          =   315
            Left            =   6750
            TabIndex        =   14
            Top             =   1395
            Visible         =   0   'False
            Width           =   495
         End
      End
      Begin MSFlexGridLib.MSFlexGrid GRDTranx 
         Height          =   6930
         Left            =   75
         TabIndex        =   3
         Top             =   1455
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12224
         _Version        =   393216
         Rows            =   1
         Cols            =   12
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         Caption         =   "TOTAL"
         ForeColor       =   &H000000FF&
         Height          =   870
         Left            =   105
         TabIndex        =   26
         Top             =   8385
         Width           =   9375
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
            Index           =   10
            Left            =   7395
            TabIndex        =   34
            Top             =   150
            Width           =   1815
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
            TabIndex        =   33
            Top             =   435
            Width           =   1830
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
            Index           =   7
            Left            =   3180
            TabIndex        =   32
            Top             =   150
            Width           =   1905
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
            TabIndex        =   31
            Top             =   435
            Width           =   1965
         End
         Begin VB.Label LBLTOTAL 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Rcvd Amt"
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
            Index           =   6
            Left            =   5340
            TabIndex        =   30
            Top             =   150
            Width           =   1875
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
            TabIndex        =   29
            Top             =   435
            Width           =   1875
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
            TabIndex        =   28
            Top             =   435
            Width           =   1965
         End
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
            Index           =   3
            Left            =   1035
            TabIndex        =   27
            Top             =   150
            Width           =   1905
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref #"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   300
         Index           =   0
         Left            =   6375
         TabIndex        =   40
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label LBLSelected 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   540
         Left            =   6345
         TabIndex        =   37
         Top             =   375
         Width           =   1980
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
         Index           =   11
         Left            =   6420
         TabIndex        =   36
         Top             =   150
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Press F6 to make Part Payments"
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
         Left            =   165
         TabIndex        =   25
         Top             =   9225
         Visible         =   0   'False
         Width           =   6300
      End
   End
End
Attribute VB_Name = "FRMRcpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer

Private Sub CMDDISPLAY_Click()
    Call Fillgrid
End Sub

Private Sub CmdExit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CmdPay_Click()
    
    Dim RSTTRXFILE, rstBILL As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim RSTITEMMAST, RSTITEMMAST2, RSTITEMMAST3 As ADODB.Recordset
    Dim i, BillNO, max, max_rpt As Long
    
    On Error GoTo eRRHAND
    If GRDTranx.Rows <= 1 Then Exit Sub
    If grdcount.Rows = 0 Then
        MsgBox "Please Select atleast one", vbOKOnly, "Payments"
        Exit Sub
    End If
    If Val(LBLSelected.Caption) = 0 Then
        MsgBox "No amount selected", vbOKOnly, "Payments"
        Exit Sub
    End If
    If DataList2.BoundText = "" Then
        MsgBox "Please select the Supplier from the list", vbOKOnly, "Payments"
        Exit Sub
    End If
    
    If Trim(TXTREFNO.Text) = "" Then
        MsgBox "Please Enter the Cheque/ Draft / Receipt No.", vbOKOnly, "Payments"
        TXTREFNO.SetFocus
        Exit Sub
    End If
    If MsgBox("Are you sure you want to make Payment of Rs. " & LBLSelected.Caption & " for the selected invoices", vbYesNo, "Payments") = vbNo Then Exit Sub
    max_rpt = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(Val(CR_NO)) From DBTPYMT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        max_rpt = IIf(IsNull(rstMaxRec.Fields(0)), 1, rstMaxRec.Fields(0) + 1)
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    max = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(Val(REC_NO)) From CASHATRXFILE ", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        max = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing
    
    For i = 0 To grdcount.Rows - 1
        If Val(grdcount.TextMatrix(i, 6)) = 0 Then GoTo SKIP
        Set rstBILL = New ADODB.Recordset
        rstBILL.Open "Select MAX(Val(RCPT_NO)) From TRNXRCPT WHERE TRX_TYPE = 'RT'", db, adOpenForwardOnly
        If Not (rstBILL.EOF And rstBILL.BOF) Then
            BillNO = IIf(IsNull(rstBILL.Fields(0)), 1, rstBILL.Fields(0) + 1)
        End If
        rstBILL.Close
        Set rstBILL = Nothing
    
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "Select * From TRNXRCPT ", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "RT"
        RSTTRXFILE!RCPT_NO = BillNO
        RSTTRXFILE!INV_NO = Val(grdcount.TextMatrix(i, 3))
        RSTTRXFILE!RCPT_DATE = Format(grdcount.TextMatrix(i, 2), "DD/MM/YYYY")
        RSTTRXFILE!RCPT_AMOUNT = Val(grdcount.TextMatrix(i, 6))
        RSTTRXFILE!ACT_CODE = DataList2.BoundText
        RSTTRXFILE!ACT_NAME = DataList2.Text
        RSTTRXFILE!RCPT_ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTTRXFILE!REF_NO = Trim(TXTREFNO.Text)
        RSTTRXFILE!INV_DATE = Format(grdcount.TextMatrix(i, 2), "DD/MM/YYYY")
        RSTTRXFILE!CR_NO = max_rpt
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
        Set RSTITEMMAST3 = New ADODB.Recordset
        RSTITEMMAST3.Open "Select * From DBTPYMT", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTITEMMAST3.AddNew
        RSTITEMMAST3!TRX_TYPE = "RT"
        RSTITEMMAST3!CR_NO = max_rpt
        'RSTITEMMAST3!INV_NO = Val(lblinvno.Caption)
        RSTITEMMAST3!RCPT_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST3!RCPT_AMT = Val(grdcount.TextMatrix(i, 6))
        RSTITEMMAST3!ACT_CODE = DataList2.BoundText
        RSTITEMMAST3!ACT_NAME = DataList2.Text
        RSTITEMMAST3!INV_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST3!REF_NO = Trim(TXTREFNO.Text)
        RSTITEMMAST3!INV_AMT = Null
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTITEMMAST3.Update
        RSTITEMMAST3.Close
        Set RSTITEMMAST3 = Nothing
    
        max_rpt = max_rpt + 1
        
        Set RSTITEMMAST2 = New ADODB.Recordset
        RSTITEMMAST2.Open "Select * From DBTPYMT WHERE INV_NO = " & Val(grdcount.TextMatrix(i, 3)) & " AND TRX_TYPE='DR' ", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTITEMMAST2.EOF And RSTITEMMAST2.BOF) Then
            RSTITEMMAST2!RCPT_AMT = RSTITEMMAST2!RCPT_AMT + Val(grdcount.TextMatrix(i, 6))
            RSTITEMMAST2!BAL_AMT = RSTITEMMAST2!INV_AMT - RSTITEMMAST2!RCPT_AMT
            If RSTITEMMAST2!BAL_AMT <= 0 Then RSTITEMMAST2!CHECK_FLAG = "Y" Else RSTITEMMAST2!CHECK_FLAG = "N"
            RSTITEMMAST2.Update
        End If
        RSTITEMMAST2.Close
        Set RSTITEMMAST2 = Nothing
        
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & BillNO & " AND INV_TYPE = 'RT' ", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!rec_no = max + 1
            RSTITEMMAST!INV_TYPE = "RT"
            RSTITEMMAST!INV_NO = BillNO
        End If
        RSTITEMMAST!TRX_TYPE = "DR"
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = DataList2.Text
        RSTITEMMAST!AMOUNT = Val(grdcount.TextMatrix(i, 6))
        RSTITEMMAST!VCH_DATE = Format(grdcount.TextMatrix(i, 2), "DD/MM/YYYY")
        RSTITEMMAST!CHECK_FLAG = "S"
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        
        max = max + 1
SKIP:
    Next i
    Call Fillgrid
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    Dim TOPROWCOUNT, CURROW, CURCOL As Integer
    
    TOPROWCOUNT = GRDTranx.TopRow
    CURROW = GRDTranx.Row
    CURCOL = GRDTranx.Col
    Call Fillgrid
    GRDTranx.TopRow = TOPROWCOUNT
    GRDTranx.Row = CURROW
    GRDTranx.Col = CURCOL
    GRDTranx.SetFocus
End Sub

Private Sub Form_Load()
    
    GRDTranx.TextMatrix(0, 0) = "STATUS"
    GRDTranx.TextMatrix(0, 1) = "SL"
    GRDTranx.TextMatrix(0, 2) = "INV DATE"
    GRDTranx.TextMatrix(0, 3) = "INV NO"
    GRDTranx.TextMatrix(0, 4) = "INV AMT"
    GRDTranx.TextMatrix(0, 5) = "RCVD AMT"
    GRDTranx.TextMatrix(0, 6) = "BAL AMT"
    GRDTranx.TextMatrix(0, 7) = "TYPE"
    GRDTranx.TextMatrix(0, 8) = "DAYS"
    GRDTranx.TextMatrix(0, 9) = "INV TYPE"
    GRDTranx.TextMatrix(0, 10) = ""
    GRDTranx.TextMatrix(0, 11) = "FLAG"
    
    GRDTranx.ColWidth(0) = 700
    GRDTranx.ColWidth(1) = 600
    GRDTranx.ColWidth(2) = 1300
    GRDTranx.ColWidth(3) = 1400
    GRDTranx.ColWidth(4) = 1100
    GRDTranx.ColWidth(5) = 1100
    GRDTranx.ColWidth(6) = 1100
    GRDTranx.ColWidth(7) = 1200
    GRDTranx.ColWidth(8) = 700
    GRDTranx.ColWidth(9) = 900
    GRDTranx.ColWidth(10) = 500
    GRDTranx.ColWidth(11) = 0
    
    GRDTranx.ColAlignment(0) = 3
    GRDTranx.ColAlignment(1) = 3
    GRDTranx.ColAlignment(2) = 3
    GRDTranx.ColAlignment(3) = 3
    GRDTranx.ColAlignment(4) = 3
    GRDTranx.ColAlignment(5) = 3
    GRDTranx.ColAlignment(6) = 3
    GRDTranx.ColAlignment(7) = 3
    GRDTranx.ColAlignment(8) = 3
    
    GRDBILL.TextMatrix(0, 0) = "SL"
    GRDBILL.TextMatrix(0, 1) = "RCPT Date"
    GRDBILL.TextMatrix(0, 2) = "RCVD Amt"
    GRDBILL.TextMatrix(0, 3) = "RCPT No"
    GRDBILL.TextMatrix(0, 4) = "Entry Date"
    GRDBILL.TextMatrix(0, 5) = "Ref No"
    
    GRDBILL.ColWidth(0) = 500
    GRDBILL.ColWidth(1) = 1200
    GRDBILL.ColWidth(2) = 1200
    GRDBILL.ColWidth(3) = 1200
    GRDBILL.ColWidth(4) = 1200
    GRDBILL.ColWidth(5) = 1200

    GRDBILL.ColAlignment(0) = 3
    GRDBILL.ColAlignment(1) = 3
    GRDBILL.ColAlignment(2) = 3
    GRDBILL.ColAlignment(3) = 3
    GRDBILL.ColAlignment(4) = 3
    GRDBILL.ColAlignment(5) = 3
    
    CLOSEALL = 1
    ACT_FLAG = True
    'Width = 9585
    'Height = 10185
    Left = 1500
    Top = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If ACT_FLAG = False Then ACT_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub GRDBILL_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            FRMEMAIN.Enabled = True
            FRMEBILL.Visible = False
            Call Fillgrid
            GRDTranx.SetFocus
    End Select
End Sub

Private Sub GRDBILL_KeyPress(KeyAscii As Integer)
    
    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    Select Case KeyAscii
        Case vbKeyD, Asc("d")
            CMDDISPLAY.Tag = KeyAscii
        Case vbKeyE, Asc("e")
            CMDEXIT.Tag = KeyAscii
        Case vbKeyL, Asc("l")
                If GRDBILL.Rows = 1 Then Exit Sub
                If Val(CMDDISPLAY.Tag) = 68 Or Val(CMDDISPLAY.Tag) = 100 Or Val(CMDEXIT.Tag) = 69 Or Val(CMDEXIT.Tag) = 101 Then
                    If MsgBox("Are you sure you want to delete this entry", vbYesNo, "DELETE !!!") = vbYes Then
                        db.Execute "delete * From TRNXRCPT WHERE TRX_TYPE='RT' AND RCPT_NO = " & GRDBILL.TextMatrix(GRDBILL.Row, 3) & " "
                        db.Execute "delete * From DBTPYMT WHERE TRX_TYPE='RT' AND CR_NO = " & Val(GRDBILL.TextMatrix(GRDBILL.Row, 6)) & " "
                        db.Execute "delete * FROM CASHATRXFILE WHERE INV_NO = " & GRDBILL.TextMatrix(GRDBILL.Row, 3) & " AND INV_TYPE = 'RT'"
                        
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "Select * From DBTPYMT WHERE INV_NO = " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE='DR' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        If Not (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
                            RSTTRXFILE!RCPT_AMT = RSTTRXFILE!RCPT_AMT - Val(GRDBILL.TextMatrix(GRDBILL.Row, 2))
                            RSTTRXFILE!BAL_AMT = RSTTRXFILE!INV_AMT + RSTTRXFILE!RCPT_AMT
                            If RSTTRXFILE!BAL_AMT <= 0 Then RSTTRXFILE!CHECK_FLAG = "Y" Else RSTTRXFILE!CHECK_FLAG = "N"
                            RSTTRXFILE.Update
                        End If
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                        
                        GRDBILL.Rows = 1
                        i = 0
                        Set RSTTRXFILE = New ADODB.Recordset
                        RSTTRXFILE.Open "Select * From TRNXRCPT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND INV_NO =  " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' ", db, adOpenStatic, adLockReadOnly, adCmdText
                        Do Until RSTTRXFILE.EOF
                            i = i + 1
                            GRDBILL.Rows = GRDBILL.Rows + 1
                            GRDBILL.FixedRows = 1
                            GRDBILL.TextMatrix(i, 0) = i
                            GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!RCPT_NO
                            GRDBILL.TextMatrix(i, 1) = Format(RSTTRXFILE!RCPT_DATE, "DD/MM/YYYY")
                            GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!RCPT_AMOUNT, "0.00")
                            GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!RCPT_ENTRY_DATE
                            GRDBILL.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
                            GRDBILL.TextMatrix(i, 6) = IIf(IsNull(RSTTRXFILE!CR_NO), "", RSTTRXFILE!CR_NO)
                            RSTTRXFILE.MoveNext
                        Loop
                        RSTTRXFILE.Close
                        Set RSTTRXFILE = Nothing
                        GRDBILL.SetFocus
                        'Call fillgrid
                    Else
                        GRDTranx.SetFocus
                    End If
                End If
        Case Else
            CMDEXIT.Tag = ""
            CMDDISPLAY.Tag = ""
    End Select
End Sub

Private Sub GRDBILL_LostFocus()
    If FRMEBILL.Visible = True Then
        Frmeperiod.Enabled = True
        FRMEBILL.Visible = False
        Call Fillgrid
        GRDTranx.SetFocus
    End If
End Sub

Private Sub GRDTranx_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If GRDTranx.Rows = 1 Then Exit Sub
    'If GRDTranx.Col <> 10 Then Exit Sub
    With GRDTranx
        If .TextMatrix(.Row, 7) <> "SR" Then
            oldx = .Col
            oldy = .Row
            .Row = oldy: .Col = 10: .CellPictureAlignment = 4
                'If GRDTranx.Col = 0 Then
                    If GRDTranx.CellPicture = picChecked Then
                        Set GRDTranx.CellPicture = picUnchecked
                        '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                        'strTextCheck = .Text
                        ' When you de-select a CheckBox, we need to strip out the #
                        'strChecked = strChecked & strTextCheck & ","
                        ' Don't forget to strip off the trailing , before passing the string
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 11) = "Y"
                        Call fillcount
                    Else
                        Set GRDTranx.CellPicture = picChecked
                        '.Col = .Col + 2
                        'strTextCheck = .Text
                        'strChecked = Replace(strChecked, strTextCheck & ",", "")
                        'Debug.Print strChecked
                        .TextMatrix(.Row, 11) = "N"
                        Call fillcount
                    End If
                'End If
            .Col = oldx
            .Row = oldy
        End If
    End With
End Sub

Private Sub GRDTranx_DblClick()
'    If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "PEND" Then Exit Sub
'            FRMPaymntreg.Enabled = False
'            FRMRCPT.LBLSUPPLIER.Caption = DataList2.Text
'            FRMRCPT.lblactcode.Caption = DataList2.BoundText
'            FRMRCPT.LBLINVDATE.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
'            FRMRCPT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
'            FRMRCPT.LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
'            FRMRCPT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
'            FRMRCPT.LBLBALAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
'            'FRMRCPT.LBLTYPE.Caption = Trim(GRDTranx.TextMatrix(GRDTranx.Row, 7))
'            FRMRCPT.Show
End Sub

Private Sub GRDTranx_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim i As Integer
    Dim RSTTRXFILE As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If GRDTranx.Rows = 1 Then Exit Sub
            LBLSUPPLIER.Caption = " " & DataList2.Text
            LBLINVDATE.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 2)
            LBLBILLNO.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 3)
            LBLBILLAMT.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 4)
            LBLPAID.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 5)
            LBLBAL.Caption = " " & GRDTranx.TextMatrix(GRDTranx.Row, 6)
            
            GRDBILL.Rows = 1
            i = 0
            Set RSTTRXFILE = New ADODB.Recordset
            RSTTRXFILE.Open "Select * From TRNXRCPT WHERE ACT_CODE = '" & DataList2.BoundText & "' AND INV_NO =  " & Val(LBLBILLNO.Caption) & " AND TRX_TYPE = 'RT' ", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE.EOF
                i = i + 1
                GRDBILL.Rows = GRDBILL.Rows + 1
                GRDBILL.FixedRows = 1
                GRDBILL.TextMatrix(i, 0) = i
                GRDBILL.TextMatrix(i, 3) = RSTTRXFILE!RCPT_NO
                GRDBILL.TextMatrix(i, 1) = Format(RSTTRXFILE!RCPT_DATE, "DD/MM/YYYY")
                GRDBILL.TextMatrix(i, 2) = Format(RSTTRXFILE!RCPT_AMOUNT, "0.00")
                GRDBILL.TextMatrix(i, 4) = RSTTRXFILE!RCPT_ENTRY_DATE
                GRDBILL.TextMatrix(i, 5) = IIf(IsNull(RSTTRXFILE!REF_NO), "", RSTTRXFILE!REF_NO)
                GRDBILL.TextMatrix(i, 6) = IIf(IsNull(RSTTRXFILE!CR_NO), "", RSTTRXFILE!CR_NO)
                RSTTRXFILE.MoveNext
            Loop
            RSTTRXFILE.Close
            Set RSTTRXFILE = Nothing
            
            FRMEBILL.Visible = True
            GRDBILL.SetFocus
        Case vbKeyF6
            Exit Sub
'            If GRDTranx.TextMatrix(GRDTranx.Row, 0) <> "PEND" Then Exit Sub
'            Me.Enabled = False
'            FRMRCPT.LBLSUPPLIER.Caption = DataList2.Text
'            FRMRCPT.lblactcode.Caption = DataList2.BoundText
'            FRMRCPT.LBLINVDATE.Caption = Format(GRDTranx.TextMatrix(GRDTranx.Row, 2), "DD/MM/YYYY")
'            FRMRCPT.lblinvno.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 3)
'            FRMRCPT.LBLBILLAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 4)
'            FRMRCPT.lblrcvdamt.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 5)
'            FRMRCPT.LBLBALAMT.Caption = GRDTranx.TextMatrix(GRDTranx.Row, 6)
'            FRMRCPT.LBLTYPE.Caption = Trim(GRDTranx.TextMatrix(GRDTranx.Row, 9))
'            FRMRCPT.Show
    End Select
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [CUSTMAST]  WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [CUSTMAST]  WHERE ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    GRDTranx.Rows = 1
    'LBL.Caption = ""
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Customer From List", vbOKOnly, "Sale Bil..."
                DataList2.SetFocus
                Exit Sub
            End If
            CMDDISPLAY.SetFocus
        Case vbKeyEscape
           
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Public Function Fillgrid()
    Dim rstTRANX As ADODB.Recordset
    Dim i As Integer
    
    
    If DataList2.BoundText = "" Then Exit Function
   ' On Error GoTo eRRhAND
    
    Screen.MousePointer = vbHourglass
    
    GRDTranx.Rows = 1
    LBLINVAMT.Caption = ""
    LBLPAIDAMT.Caption = ""
    LBLBALAMT.Caption = ""
    lblOPBal.Caption = ""
    i = 1
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From DBTPYMT WHERE [ACT_CODE] = '" & DataList2.BoundText & "' AND (TRX_TYPE='DR' or TRX_TYPE='SR') ORDER BY INV_DATE DESC, INV_NO DESC", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rstTRANX.EOF
        GRDTranx.Visible = False
        GRDTranx.Rows = GRDTranx.Rows + 1
        GRDTranx.FixedRows = 1
        GRDTranx.TextMatrix(i, 1) = i
        GRDTranx.TextMatrix(i, 2) = Format(rstTRANX!INV_DATE, "DD/MM/YYYY")
        GRDTranx.TextMatrix(i, 3) = IIf(IsNull(rstTRANX!INV_NO), "", rstTRANX!INV_NO)
        GRDTranx.TextMatrix(i, 4) = Format(rstTRANX!INV_AMT, "0.00")
        GRDTranx.TextMatrix(i, 7) = rstTRANX!TRX_TYPE
        Select Case rstTRANX!TRX_TYPE
            Case "DR"
                If rstTRANX!CHECK_FLAG = "Y" Then
                    GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!INV_AMT, "0.00")
                    GRDTranx.TextMatrix(i, 0) = "PAID"
                Else
                    GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                    GRDTranx.TextMatrix(i, 0) = "PEND"
                    GRDTranx.TextMatrix(i, 8) = DateDiff("d", GRDTranx.TextMatrix(i, 2), Date)
                End If
            Case Else
                GRDTranx.TextMatrix(i, 5) = Format(rstTRANX!RCPT_AMT, "0.00")
                GRDTranx.TextMatrix(i, 0) = "SALES RETURN"
        End Select
        GRDTranx.TextMatrix(i, 6) = Format(rstTRANX!INV_AMT - Val(GRDTranx.TextMatrix(i, 5)), "0.00")
        GRDTranx.TextMatrix(i, 9) = ""
        GRDTranx.TextMatrix(i, 11) = "N"
        With GRDTranx
            If .TextMatrix(i, 7) <> "SR" Then
                .Row = i: .Col = 10: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            End If
        End With
        
        GRDTranx.Row = i
        GRDTranx.Col = 0
        If rstTRANX!CHECK_FLAG = "N" Then
            LBLBALAMT.Caption = Format(Val(LBLBALAMT.Caption) + Val(GRDTranx.TextMatrix(i, 6)), "0.00")
            LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + rstTRANX!RCPT_AMT, "0.00")
            GRDTranx.CellForeColor = vbRed
        Else
            LBLPAIDAMT.Caption = Format(Val(LBLPAIDAMT.Caption) + rstTRANX!INV_AMT, "0.00")
            GRDTranx.CellForeColor = vbBlue
        End If
        LBLINVAMT.Caption = Format(Val(LBLINVAMT.Caption) + rstTRANX!INV_AMT, "0.00")
        LBLBALAMT.Caption = Val(LBLINVAMT.Caption) - Val(LBLPAIDAMT.Caption)
        i = i + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    GRDTranx.Visible = True
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "select OPEN_DB from [CUSTMAST]  WHERE ACT_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (rstTRANX.EOF And rstTRANX.BOF) Then
        lblOPBal.Caption = IIf(IsNull(rstTRANX!OPEN_DB), 0, Format(rstTRANX!OPEN_DB, "0.00"))
    End If
    rstTRANX.Close
    Set rstTRANX = Nothing
    LBLBALAMT.Caption = Format(Val(LBLBALAMT.Caption) + Val(lblOPBal.Caption), "0.00")
        
    TXTREFNO.Text = ""
    flagchange.Caption = ""
    GRDTranx.SetFocus
    Screen.MousePointer = vbDefault
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description
End Function

Private Function fillcount()
    Dim i, N As Long
    
    grdcount.Rows = 0
    i = 0
    LBLSelected.Caption = ""
    On Error GoTo eRRHAND
    For N = 1 To GRDTranx.Rows - 1
        If GRDTranx.TextMatrix(N, 11) = "Y" Then
            grdcount.Rows = grdcount.Rows + 1
            grdcount.TextMatrix(i, 0) = GRDTranx.TextMatrix(N, 0)
            grdcount.TextMatrix(i, 1) = GRDTranx.TextMatrix(N, 1)
            grdcount.TextMatrix(i, 2) = GRDTranx.TextMatrix(N, 2)
            grdcount.TextMatrix(i, 3) = GRDTranx.TextMatrix(N, 3)
            grdcount.TextMatrix(i, 4) = GRDTranx.TextMatrix(N, 4)
            grdcount.TextMatrix(i, 5) = GRDTranx.TextMatrix(N, 5)
            grdcount.TextMatrix(i, 6) = GRDTranx.TextMatrix(N, 6)
            grdcount.TextMatrix(i, 7) = GRDTranx.TextMatrix(N, 7)
            grdcount.TextMatrix(i, 8) = GRDTranx.TextMatrix(N, 8)
            grdcount.TextMatrix(i, 9) = GRDTranx.TextMatrix(N, 9)
            
            LBLSelected.Caption = Val(LBLSelected.Caption) + Val(GRDTranx.TextMatrix(N, 6))
            i = i + 1
        End If
    Next N
    
    LBLSelected.Caption = Format(LBLSelected.Caption, "0.00")
    Exit Function
eRRHAND:
    MsgBox Err.Description
    
End Function

'Public Function Refreshgrid()
'    Dim N As Long
'    Dim oldx, oldy As Variant
'    On Error GoTo eRRHAND
'    For N = 1 To GRDTranx.Rows - 1
'        GRDTranx.TextMatrix(N, 11) = "N"
'        With GRDTranx
'            oldx = 10
'            oldy = N
'            .Row = oldy: .Col = 10: .CellPictureAlignment = 4
'            Set GRDTranx.CellPicture = picChecked
'        End With
'    Next N
'
'    If grdcount.TextMatrix(0, 1) = "" Then GoTo SKIP
'    'GRDTranx.TextMatrix (grdcount.TextMatrix(n, 1))
'    For N = 0 To grdcount.Rows - 1
'        GRDTranx.TextMatrix(grdcount.TextMatrix(N, 1), 11) = "Y"
'        With GRDTranx
'            oldx = 10
'            oldy = grdcount.TextMatrix(N, 1)
'            .Row = oldy: .Col = 10: .CellPictureAlignment = 4
'            Set GRDTranx.CellPicture = picUnchecked
'        End With
'    Next N
'
'SKIP:
'    Call fillcount
'    Exit Function
'eRRHAND:
'    MsgBox Err.Description
'
'End Function

