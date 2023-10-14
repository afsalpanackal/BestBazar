VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmFixedAssets 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIXED ASSETS ENTRY"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11985
   ControlBox      =   0   'False
   Icon            =   "frmfxdassets.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   11985
   Begin VB.CommandButton Command4 
      Caption         =   "<<&Previous"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6555
      TabIndex        =   42
      Top             =   6765
      Width           =   1170
   End
   Begin VB.CommandButton Command5 
      Caption         =   "&Next>>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6555
      TabIndex        =   41
      Top             =   7185
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   8730
      TabIndex        =   32
      Top             =   1545
      Width           =   3270
      Begin VB.Label lblmode 
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   2580
         TabIndex        =   43
         Top             =   225
         Width           =   480
      End
      Begin VB.Label lblremarks 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1050
         TabIndex        =   39
         Top             =   1965
         Width           =   2490
      End
      Begin VB.Label lblcash 
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   45
         TabIndex        =   38
         Top             =   2025
         Width           =   975
      End
      Begin VB.Label lblbankname 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   45
         TabIndex        =   37
         Top             =   1635
         Width           =   2490
      End
      Begin VB.Label lblbankcode 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   30
         TabIndex        =   36
         Top             =   1275
         Width           =   2490
      End
      Begin VB.Label lblpassflag 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   45
         TabIndex        =   35
         Top             =   915
         Width           =   2490
      End
      Begin VB.Label lblchqdate 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   45
         TabIndex        =   34
         Top             =   540
         Width           =   2490
      End
      Begin VB.Label lblchqno 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   60
         TabIndex        =   33
         Top             =   210
         Width           =   2490
      End
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
      TabIndex        =   6
      Top             =   7635
      Width           =   915
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
      Left            =   1110
      TabIndex        =   21
      Top             =   285
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
      TabIndex        =   7
      Top             =   7635
      Width           =   1065
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00F4DFCE&
      Height          =   8235
      Left            =   -120
      TabIndex        =   8
      Top             =   -45
      Width           =   7920
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00F4DFCE&
         Height          =   585
         Left            =   150
         TabIndex        =   13
         Top             =   150
         Width           =   7665
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
            Left            =   6825
            TabIndex        =   19
            Top             =   225
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
            Left            =   3045
            MaxLength       =   10
            TabIndex        =   17
            Top             =   195
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   4935
            TabIndex        =   18
            Top             =   195
            Width           =   1395
            _ExtentX        =   2461
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
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   0
            TabIndex        =   29
            Top             =   480
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   0
            TabIndex        =   28
            Top             =   390
            Visible         =   0   'False
            Width           =   495
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
            Left            =   6420
            TabIndex        =   20
            Top             =   210
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry Date"
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
            Left            =   1995
            TabIndex        =   16
            Top             =   180
            Width           =   1050
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Entry No."
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
            TabIndex        =   15
            Top             =   180
            Width           =   885
         End
         Begin VB.Label INVDATE 
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            Left            =   4410
            TabIndex        =   14
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00F4DFCE&
         Height          =   2115
         Left            =   135
         TabIndex        =   9
         Top             =   6045
         Width           =   7755
         Begin VB.CommandButton CmdPrint 
            Caption         =   "&Print"
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
            Left            =   6675
            TabIndex        =   40
            Top             =   1635
            Width           =   1020
         End
         Begin VB.TextBox txtremarks 
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
            MaxLength       =   30
            TabIndex        =   26
            Top             =   450
            Width           =   2040
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   720
            TabIndex        =   22
            Top             =   450
            Width           =   3690
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
            TabIndex        =   2
            Top             =   1635
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
            Left            =   45
            TabIndex        =   0
            Top             =   450
            Width           =   645
         End
         Begin VB.TextBox TXTAMOUNT 
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
            Left            =   4455
            MaxLength       =   7
            TabIndex        =   1
            Top             =   450
            Width           =   1170
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
            TabIndex        =   4
            Top             =   1635
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
            TabIndex        =   3
            Top             =   1635
            Width           =   1155
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
            TabIndex        =   5
            Top             =   1635
            Width           =   960
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   840
            Left            =   735
            TabIndex        =   23
            Top             =   765
            Width           =   3675
            _ExtentX        =   6482
            _ExtentY        =   1482
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
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "TOTAL AMOUNT"
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
            Height          =   225
            Index           =   6
            Left            =   4500
            TabIndex        =   31
            Top             =   900
            Width           =   1620
         End
         Begin VB.Label lbltotal 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800080&
            Height          =   435
            Left            =   4470
            TabIndex        =   30
            Top             =   1140
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Remarks"
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
            Index           =   2
            Left            =   5655
            TabIndex        =   27
            Top             =   195
            Width           =   2040
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Assets Head"
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
            Left            =   720
            TabIndex        =   24
            Top             =   195
            Width           =   3690
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "SL"
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
            Left            =   45
            TabIndex        =   12
            Top             =   195
            Width           =   645
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Amount"
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
            Left            =   4455
            TabIndex        =   11
            Top             =   195
            Width           =   1170
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
            TabIndex        =   10
            Top             =   1770
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   5265
         Left            =   150
         TabIndex        =   25
         Top             =   780
         Width           =   7695
         _ExtentX        =   13573
         _ExtentY        =   9287
         _Version        =   393216
         Rows            =   1
         Cols            =   5
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         HighLight       =   0
         AllowUserResizing=   1
         Appearance      =   0
         GridLineWidth   =   2
      End
   End
End
Attribute VB_Name = "frmFixedAssets"
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

Private Sub CMDADD_Click()
    Dim i As Long
    
    If grdsales.rows <= Val(TXTSLNO.Text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = DataList2.BoundText
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(DataList2.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Format(Val(TXTAMOUNT.Text), "0.00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Trim(TXTREMARKS.Text)
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.rows - 1
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 3))
    Next i
    
    TXTSLNO.Text = grdsales.rows
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    TXTREMARKS.Text = ""
    cmdRefresh.Enabled = True
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    txtBillNo.Enabled = False
    
    If grdsales.rows >= 18 Then grdsales.TopRow = grdsales.rows - 1

End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdadd.Enabled = False
            TXTAMOUNT.Enabled = True
            TXTAMOUNT.SetFocus
            Exit Sub
    End Select

End Sub

Private Sub cmdcancel_Click()
    If M_ADD = True Then
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo, "Fixed Assets") = vbNo Then Exit Sub
    End If
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTSLNO.Text = ""
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    TXTREMARKS = ""
    lblcash.Caption = "Y"
    lblbankcode.Caption = ""
    lblbankname.Caption = ""
    lblchqdate.Caption = ""
    lblchqno.Caption = ""
    lblremarks.Caption = ""
    lblmode.Caption = ""
    lblpassflag.Caption = ""
    LBLTOTAL.Caption = "0.00"
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETS", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
'        TXTLASTBILL.Text = txtBillNo.Text
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CmdDelete_Click()
    Dim i As Long
    
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & grdsales.TextMatrix(Val(TXTSLNO.Text), 2) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
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
    
    For i = Val(TXTSLNO.Text) To grdsales.rows - 2
        'grdsales.TextMatrix(Val(TXTSLNO.text), 0) = grdsales.TextMatrix(i + 1, 0)
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
    Next i
    grdsales.rows = grdsales.rows - 1
    TXTSLNO.Text = Val(grdsales.rows)
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 3))
    Next i
    
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    TXTREMARKS = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    cmdRefresh.Enabled = True
    M_ADD = True
    If grdsales.rows = 1 Then
        CMDEXIT.Enabled = True
    End If
End Sub

Private Sub CmdExit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.Text) >= grdsales.rows Then Exit Sub
    
    M_EDIT = True
    CMDMODIFY.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    cmdRefresh.Enabled = True
    TXTAMOUNT.Enabled = True
    TXTAMOUNT.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.rows
            TXTDEALER.Text = ""
            TXTAMOUNT.Text = ""
            TXTREMARKS.Text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTDEALER.Enabled = False
            DataList2.Enabled = False
            TXTAMOUNT.Enabled = False
            TXTREMARKS.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub CmdPrint_Click()
    Dim i As Long
    On Error GoTo ERRHAND
    If grdsales.rows = 1 Then
        'MsgBox "Please Select Purchase Order No.", vbOKOnly, "Print Voucher..."
        Exit Sub
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
    
    ReportNameVar = Rptpath & "RPTExpense"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    
    Report.RecordSelectionFormula = "({TRXFXDASSETMAST.TRX_TYPE} ='FA' AND {TRXFXDASSETMAST.VCH_NO}= " & Val(txtBillNo.Text) & " AND {TRXFXDASSETMAST.TRX_YEAR}= '" & Year(MDIMAIN.DTFROM.Value) & "' )"
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
        If CRXFormulaField.Name = "{@state}" Then CRXFormulaField.Text = "'" & "State Code: " & Trim(MDIMAIN.LBLSTATE.Caption) & "(" & Trim(MDIMAIN.LBLSTATENAME.Caption) & ")" & "'"
        If CRXFormulaField.Name = "{@Comp_Name}" Then CRXFormulaField.Text = "'" & CompName & "'"
        If CRXFormulaField.Name = "{@Comp_Address1}" Then CRXFormulaField.Text = "'" & CompAddress1 & "'"
        If CRXFormulaField.Name = "{@Comp_Address2}" Then CRXFormulaField.Text = "'" & CompAddress2 & "'"
        If CRXFormulaField.Name = "{@Comp_Address3}" Then CRXFormulaField.Text = "'" & CompAddress3 & "'"
        If CRXFormulaField.Name = "{@Comp_Address4}" Then CRXFormulaField.Text = "'" & CompAddress4 & "'"
        If CRXFormulaField.Name = "{@Comp_Tin}" Then CRXFormulaField.Text = "'" & CompTin & "'"
        If CRXFormulaField.Name = "{@Comp_CST}" Then CRXFormulaField.Text = "'" & CompCST & "'"
    Next
    frmreport.Caption = "Print Voucher"
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Sub

Private Sub cmdRefresh_Click()
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Enter the Date", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    
    If (DateValue(TXTINVDATE.Text) < DateValue(MDIMAIN.DTFROM.Value)) Or (DateValue(TXTINVDATE.Text) >= DateValue(DateAdd("YYYY", 1, MDIMAIN.DTFROM.Value))) Then
        MsgBox "Please check the Date", vbOKOnly, "EzBiz"
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    Set creditbill = Me
    Me.Enabled = False
    FRMCSBK.Show
    Screen.MousePointer = vbNormal
    'Call appendpurchase
    Exit Sub
    
    
End Sub

Private Sub Command4_Click()
    If Val(txtBillNo.Text) <= 1 Then Exit Sub
    If M_ADD = True Then
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo, "Fixed Assets") = vbNo Then Exit Sub
    End If
    txtBillNo.Enabled = True
    'txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTSLNO.Text = ""
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    TXTREMARKS = ""
    lblcash.Caption = "Y"
    lblbankcode.Caption = ""
    lblbankname.Caption = ""
    lblchqdate.Caption = ""
    lblchqno.Caption = ""
    lblremarks.Caption = ""
    lblmode.Caption = ""
    lblpassflag.Caption = ""
    LBLTOTAL.Caption = "0.00"
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    
    txtBillNo.Text = Val(txtBillNo.Text) - 1
    Call txtBillNo_KeyDown(13, 0)
    TXTINVDATE.SetFocus
'    Dim TRXMAST As ADODB.Recordset
'    On Error GoTo Errhand
'
'
'
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
'        TXTLASTBILL.Text = txtBillNo.Text
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Command5_Click()
    If M_ADD = True Then
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo, "Fixed Assets") = vbNo Then Exit Sub
    End If
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND

    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Tag = IIf(IsNull(TRXMAST.Fields(0)), 0, TRXMAST.Fields(0))
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    If Val(txtBillNo.Text) > Val(txtBillNo.Tag) Then Exit Sub
    
    txtBillNo.Enabled = True
    'txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = "  /  /    "
    TXTSLNO.Text = ""
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    TXTREMARKS = ""
    lblcash.Caption = "Y"
    lblbankcode.Caption = ""
    lblbankname.Caption = ""
    lblchqdate.Caption = ""
    lblchqno.Caption = ""
    lblremarks.Caption = ""
    lblmode.Caption = ""
    lblpassflag.Caption = ""
    LBLTOTAL.Caption = "0.00"
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    
    txtBillNo.Text = Val(txtBillNo.Text) + 1
    Call txtBillNo_KeyDown(13, 0)
    TXTINVDATE.SetFocus
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Form_Activate()
    On Error GoTo ERRHAND
    txtBillNo.SetFocus
    Exit Sub
ERRHAND:
    If err.Number = 5 Then Exit Sub
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETS", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
'        TXTLASTBILL.Text = txtBillNo.Text
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    lblcash.Caption = "Y"
    lblbankcode.Caption = ""
    lblbankname.Caption = ""
    lblchqdate.Caption = ""
    lblchqno.Caption = ""
    lblremarks.Caption = ""
    lblmode.Caption = ""
    lblpassflag.Caption = ""
    grdsales.ColWidth(0) = 400
    grdsales.ColWidth(1) = 1000
    grdsales.ColWidth(2) = 3500
    grdsales.ColWidth(3) = 1000
    grdsales.ColWidth(4) = 2500
    
    
    grdsales.ColAlignment(0) = 4
    grdsales.ColAlignment(1) = 1
    grdsales.ColAlignment(2) = 1
    grdsales.ColAlignment(3) = 7
    grdsales.ColAlignment(4) = 1
    
    grdsales.TextArray(0) = "SL"
    grdsales.TextArray(1) = "ACT_CODE"
    grdsales.TextArray(2) = "ASSETS HEAD"
    grdsales.TextArray(3) = "AMOUNT"
    grdsales.TextArray(4) = "REMARKS"

    PHYFLAG = True
    TXTDEALER.Enabled = False
    DataList2.Enabled = False
    TXTAMOUNT.Enabled = False
    TXTINVDATE.Text = Date
    TXTDATE.Text = Date
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    TXTSLNO.Text = 1
    TXTSLNO.Enabled = True
    FRMECONTROLS.Enabled = False
    FRMEMASTER.Enabled = False
    CLOSEALL = 1
    M_ADD = False
    Me.Width = 7935
    Me.Height = 9435
    Me.Left = 0
    Me.Top = 0
    LBLTOTAL.Caption = "0.00"
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ACT_FLAG = False Then ACT_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub txtBillNo_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub txtBillNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ERRHAND
    LBLTOTAL.Caption = ""
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.rows = 1
            i = 0
            grdsales.rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRXFXDASSETS WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE ='FA' ORDER BY VCH_NO, VCH_DATE", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ACT_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ACT_NAME
                grdsales.TextMatrix(i, 3) = Format(rstTRXMAST!VCH_AMOUNT, "0.00")
                grdsales.TextMatrix(i, 4) = IIf(IsNull(rstTRXMAST!REMARKS), "", rstTRXMAST!REMARKS)
                LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + rstTRXMAST!VCH_AMOUNT, "0.00")
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE ='FA' ORDER BY VCH_NO, VCH_DATE", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                lblcash.Caption = IIf(IsNull(rstTRXMAST!Cash_Flag), "Y", rstTRXMAST!Cash_Flag)
                lblbankcode.Caption = IIf(IsNull(rstTRXMAST!BANK_CODE), "", rstTRXMAST!BANK_CODE)
                lblbankname.Caption = IIf(IsNull(rstTRXMAST!BANK_NAME), "", rstTRXMAST!BANK_NAME)
                lblremarks.Caption = IIf(IsNull(rstTRXMAST!BANK_REMARKS), "", rstTRXMAST!BANK_REMARKS)
                lblmode.Caption = IIf(IsNull(rstTRXMAST!BANK_MODE), "", rstTRXMAST!BANK_MODE)
                lblchqdate.Caption = IIf(IsDate(rstTRXMAST!BANK_CHQ_DATE), Format(rstTRXMAST!BANK_CHQ_DATE, "DD/MM/YYYY"), "")
                lblchqno.Caption = IIf(IsNull(rstTRXMAST!BANK_CHQ_NO), "", rstTRXMAST!BANK_CHQ_NO)
                lblpassflag.Caption = IIf(IsNull(rstTRXMAST!BANK_CHQ_FLAG), "", rstTRXMAST!BANK_CHQ_FLAG)
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            TXTSLNO.Text = grdsales.rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If i > 0 Or (Val(txtBillNo.Text) < Val(TXTLASTBILL.Text)) Then
                'FRMEMASTER.Enabled = False
                cmdcancel.SetFocus
            Else
                TXTINVDATE.SetFocus
            End If
    End Select
    
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
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
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
    M_EDIT = False
End Sub

Private Sub TXTINVDATE_GotFocus()
    TXTINVDATE.SelStart = 0
    TXTINVDATE.SelLength = Len(TXTINVDATE.Text)
End Sub

Private Sub TXTINVDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTINVDATE.Text = "  /  /    " Then
                TXTINVDATE.Text = Format(Date, "DD/MM/YYYY")
                FRMECONTROLS.Enabled = True
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
                Exit Sub
            End If
            If Not IsDate(TXTINVDATE.Text) Then
                TXTINVDATE.SetFocus
            Else
                TXTINVDATE.Text = Format(TXTINVDATE.Text, "DD/MM/YYYY")
                FRMECONTROLS.Enabled = True
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            End If
        Case vbKeyEscape
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

Private Sub TXTAMOUNT_GotFocus()
    TXTAMOUNT.SelStart = 0
    TXTAMOUNT.SelLength = Len(TXTAMOUNT.Text)
End Sub

Private Sub TXTAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTAMOUNT.Text) = 0 Then Exit Sub
            TXTREMARKS.Enabled = True
            TXTAMOUNT.Enabled = False
            TXTREMARKS.SetFocus
        Case vbKeyEscape
            TXTAMOUNT.Enabled = False
            TXTDEALER.Enabled = True
            DataList2.Enabled = True
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub TXTAMOUNT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTAMOUNT_LostFocus()
    TXTAMOUNT.Text = Format(TXTAMOUNT.Text, ".00")
End Sub

Private Sub TXTSLNO_GotFocus()
    TXTSLNO.SelStart = 0
    TXTSLNO.SelLength = Len(TXTSLNO.Text)
End Sub

Private Sub TXTSLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If Val(TXTSLNO.Text) = 0 Then
                TXTSLNO.Text = grdsales.rows
                CmdDelete.Enabled = False
                GoTo SKIP
            End If
            If Val(TXTSLNO.Text) >= grdsales.rows Then
                TXTSLNO.Text = grdsales.rows
                CmdDelete.Enabled = False
                CMDMODIFY.Enabled = False
            End If
            If Val(TXTSLNO.Text) < grdsales.rows Then
                TXTSLNO.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 0)
                'TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTDEALER.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                DataList2.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                Call DataList2_Click
                TXTAMOUNT.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                TXTREMARKS.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
                
                TXTSLNO.Enabled = False
                TXTDEALER.Enabled = False
                TXTAMOUNT.Enabled = False
                CMDMODIFY.Enabled = True
                CMDMODIFY.SetFocus
                CmdDelete.Enabled = True
                Exit Sub
            End If
SKIP:
            TXTSLNO.Enabled = False
            TXTDEALER.Enabled = True
            DataList2.Enabled = True
            TXTAMOUNT.Enabled = False
            TXTDEALER.SetFocus
        Case vbKeyEscape
            cmdRefresh.Enabled = True
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.rows)
                TXTDEALER.Text = ""
                TXTAMOUNT.Text = ""
                TXTREMARKS.Text = ""
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

Public Sub appendpurchase()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    db.BeginTrans
    db.Execute "delete From TRXFXDASSETS WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE ='FA'"
    db.Execute "delete From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND VCH_NO = " & Val(txtBillNo.Text) & " AND TRX_TYPE ='FA'"
    db.Execute "delete FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'FA' AND INV_TRX_TYPE = 'FA'"
    'db.Execute "delete From BANK_TRX WHERE B_TRX_YEAR='" & Year(MDIMAIN.DTFROM.value) & "' AND B_VCH_NO = " & Val(txtBillNo.Text) & " AND B_TRX_TYPE = 'GI' "
    db.Execute "delete From BANK_TRX WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE ='DR' AND BILL_TRX_TYPE = 'FA' AND B_VCH_NO = " & Val(txtBillNo.Text) & " "
    
    If grdsales.rows = 1 Then GoTo SKIP
        
    i = 0
    If lblcash.Caption = "Y" Then
        Set rstMaxRec = New ADODB.Recordset
        rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
        If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
            i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
        End If
        rstMaxRec.Close
        Set rstMaxRec = Nothing
        
'        Set RSTITEMMAST = New ADODB.Recordset
'        RSTITEMMAST.Open "Select * From CASHATRXFILE WHERE REC_NO= (SELECT MAX(REC_NO) FROM CASHATRXFILE)", db, adOpenStatic, adLockOptimistic, adCmdText
'        If RSTITEMMAST.RecordCount = 0 Then
'            i = 1
'        Else
'            i = RSTITEMMAST!REC_NO + 1
'        End If
        
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND INV_NO = " & Val(txtBillNo.Text) & " AND INV_TYPE = 'FA' AND INV_TRX_TYPE = 'FA'", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
            RSTITEMMAST.AddNew
            RSTITEMMAST!REC_NO = i + 1
            RSTITEMMAST!INV_TYPE = "FA"
            RSTITEMMAST!INV_TRX_TYPE = "FA"
            RSTITEMMAST!INV_NO = Val(txtBillNo.Text)
            RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        End If
        RSTITEMMAST!TRX_TYPE = "DR"
        RSTITEMMAST!ACT_CODE = DataList2.BoundText
        RSTITEMMAST!ACT_NAME = "Fixed Assets" 'Trim(DataList2.Text)
        RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
        RSTITEMMAST!check_flag = "P"
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    Else
        Set RSTITEMMAST = New ADODB.Recordset
        RSTITEMMAST.Open "Select * From BANK_TRX WHERE TRX_NO= (SELECT MAX(TRX_NO) FROM BANK_TRX WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "' AND TRX_TYPE = 'DR' AND BILL_TRX_TYPE = 'FA')", db, adOpenStatic, adLockOptimistic, adCmdText
        If RSTITEMMAST.RecordCount = 0 Then
            i = 1
        Else
            i = RSTITEMMAST!TRX_NO + 1
        End If
        RSTITEMMAST.AddNew
        RSTITEMMAST!TRX_TYPE = "DR"
        RSTITEMMAST!TRX_NO = i
        RSTITEMMAST!BILL_TRX_TYPE = "FA"
        RSTITEMMAST!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!B_TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTITEMMAST!B_VCH_NO = Val(txtBillNo.Text)
        RSTITEMMAST!B_TRX_TYPE = "FA"
        RSTITEMMAST!BANK_CODE = lblbankcode.Caption
        RSTITEMMAST!BANK_NAME = lblbankname.Caption
        'RSTTRXFILE!INV_NO = Val(lblinvno.Caption)
        RSTITEMMAST!TRX_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!TRX_AMOUNT = Val(LBLTOTAL.Caption)
        RSTITEMMAST!ACT_CODE = ""
        RSTITEMMAST!ACT_NAME = "Fixed Assets"
        'RSTTRXFILE!INV_DATE = LBLDATE.Caption
        RSTITEMMAST!REF_NO = Trim(lblremarks.Caption)
        RSTITEMMAST!BANK_MODE = Trim(lblmode.Caption)
        RSTITEMMAST!ENTRY_DATE = Format(TXTDATE.Text, "DD/MM/YYYY")
        RSTITEMMAST!CHQ_DATE = Format(lblchqdate.Caption, "DD/MM/YYYY")
        RSTITEMMAST!BANK_FLAG = "Y"
        If lblpassflag.Caption = "N" Then
            RSTITEMMAST!check_flag = "N"
        Else
            RSTITEMMAST!check_flag = "Y"
        End If
        RSTITEMMAST!CHQ_NO = Trim(lblchqno.Caption)
        'RSTTRXFILE!INV_DATE = Format(lblinvdate.Caption, "DD/MM/YYYY")
        RSTITEMMAST.Update
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
    End If
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from TRXFXDASSETMAST ", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "FA"
    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
    RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = ""
    RSTTRXFILE!ACT_NAME = ""
    RSTTRXFILE!Cash_Flag = lblcash.Caption
    RSTTRXFILE!BANK_CODE = lblbankcode.Caption
    RSTTRXFILE!BANK_NAME = lblbankname.Caption
    RSTTRXFILE!BANK_REMARKS = lblremarks.Caption
    RSTTRXFILE!BANK_MODE = Trim(lblmode.Caption)
    RSTTRXFILE!BANK_CHQ_DATE = lblchqdate.Caption
    RSTTRXFILE!BANK_CHQ_NO = lblchqno.Caption
    RSTTRXFILE!BANK_CHQ_FLAG = lblpassflag.Caption
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!MODIFY_DATE = Format(Date, "dd/mm/yyyy")
    RSTTRXFILE!M_USER_ID = "SM"
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    For i = 1 To grdsales.rows - 1
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT * from TRXFXDASSETS ", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "FA"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!LINE_NO = i
        RSTTRXFILE!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!ACT_CODE = Trim(grdsales.TextMatrix(i, 1))
        RSTTRXFILE!ACT_NAME = Trim(grdsales.TextMatrix(i, 2))
        RSTTRXFILE!VCH_AMOUNT = Val(grdsales.TextMatrix(i, 3))
        RSTTRXFILE!REMARKS = Trim(grdsales.TextMatrix(i, 4))
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = DataList2.BoundText
        RSTTRXFILE.Update
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing

    Next i
    
SKIP:
    db.CommitTrans
    Screen.MousePointer = vbNormal
    M_ADD = False
    CMDEXIT.Enabled = True
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "FIXED ASSETS"
    Exit Sub
    
    Dim rstMaxNo As ADODB.Recordset
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From TRXFXDASSETMAST WHERE TRX_YEAR='" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    lblcash.Caption = "Y"
    lblbankcode.Caption = ""
    lblbankname.Caption = ""
    lblchqdate.Caption = ""
    lblchqno.Caption = ""
    lblremarks.Caption = ""
    lblmode.Caption = ""
    lblpassflag.Caption = ""
    grdsales.rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    lbldealer.Caption = ""
    flagchange.Caption = ""
    TXTDEALER.Text = ""
    TXTINVDATE.Text = "  /  /    "
    TXTSLNO.Text = ""
    TXTAMOUNT.Text = ""
    LBLTOTAL.Caption = "0.00"
    TXTREMARKS.Text = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    Screen.MousePointer = vbNormal
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "FIXED ASSETS"
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    'If Err.Number <> -2147168237 Then
        MsgBox err.Description
    'End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ERRHAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ITEM_CODE, ITEM_NAME from ASTMAST  WHERE ITEM_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ITEM_CODE, ITEM_NAME from ASTMAST  WHERE ITEM_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ITEM_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "ITEM_NAME"
        DataList2.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
ERRHAND:
    MsgBox err.Description
    
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.Text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TXTDEALER.Text) = "" Then Exit Sub
            If DataList2.VisibleCount = 0 Then
                If MsgBox("Assets head not found. Do you want to create it?", vbYesNo + vbDefaultButton2, "Assets Head Creation") = vbNo Then Exit Sub
                
                On Error GoTo ERRHAND
                Dim RSTITEMMAST As ADODB.Recordset
                Dim expcode As String
                
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ASTMAST ", db, adOpenStatic, adLockReadOnly
                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    expcode = IIf(IsNull(RSTITEMMAST.Fields(0)), "1", Val(RSTITEMMAST.Fields(0)) + 1)
                End If
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ASTMAST WHERE ITEM_CODE = '" & expcode & "'", db, adOpenStatic, adLockOptimistic, adCmdText
                If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    RSTITEMMAST.AddNew
                    RSTITEMMAST!ITEM_CODE = expcode
                End If
                RSTITEMMAST!ITEM_NAME = Trim(TXTDEALER.Text)
                RSTITEMMAST!Category = "SERVICE CHARGE"
                RSTITEMMAST!UNIT = 1
                RSTITEMMAST!MANUFACTURER = "GENERAL"
                RSTITEMMAST!REMARKS = ""
                RSTITEMMAST!PACK_TYPE = ""
                RSTITEMMAST!PTR = 0
                RSTITEMMAST!OPEN_QTY = 0
                RSTITEMMAST!OPEN_VAL = 0
                RSTITEMMAST!RCPT_QTY = 0
                RSTITEMMAST!RCPT_VAL = 0
                RSTITEMMAST!ISSUE_QTY = 0
                RSTITEMMAST!ISSUE_VAL = 0
                RSTITEMMAST!CLOSE_QTY = 0
                RSTITEMMAST!CLOSE_VAL = 0
                RSTITEMMAST!DISC = 0
                RSTITEMMAST!SALES_TAX = 0
                RSTITEMMAST!check_flag = "V"
                RSTITEMMAST!item_COST = 0
                
                RSTITEMMAST.Update
                RSTITEMMAST.Close
                Set RSTITEMMAST = Nothing
                
                Call TXTDEALER_Change
                
                Exit Sub
            End If
            DataList2.SetFocus
        Case vbKeyEscape
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
    End Select
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
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

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.Text = lbldealer.Caption
    DataList2.Text = TXTDEALER.Text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub DataList2_Click()
    TXTDEALER.Text = DataList2.Text
    lbldealer.Caption = TXTDEALER.Text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.Text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Fixed Assets head From List", vbOKOnly, "Fixed Assets Entry..."
                DataList2.SetFocus
                Exit Sub
            End If
            TXTDEALER.Enabled = False
            DataList2.Enabled = False
            TXTAMOUNT.Enabled = True
            TXTAMOUNT.SetFocus
            'FRMEHEAD.Enabled = False
            'TXTSLNO.Enabled = True
            'TXTSLNO.SetFocus
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

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTREMARKS.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTAMOUNT.Enabled = True
            TXTREMARKS.Enabled = False
            TXTAMOUNT.SetFocus
    End Select
End Sub

Private Sub TXTREMARKS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREMARKS_LostFocus()
    TXTREMARKS.Text = Format(TXTREMARKS.Text, ".00")
End Sub


