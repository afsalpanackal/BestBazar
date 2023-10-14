VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmIncome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OFFICE INCOME ENTRY"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   ControlBox      =   0   'False
   Icon            =   "frmincome.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   7845
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
      Left            =   4410
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
      TabIndex        =   22
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
      Left            =   5370
      TabIndex        =   7
      Top             =   7650
      Width           =   1065
   End
   Begin VB.Frame Fram 
      BackColor       =   &H00C0E0FF&
      Height          =   8235
      Left            =   -120
      TabIndex        =   9
      Top             =   -45
      Width           =   7920
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0E0FF&
         Height          =   585
         Left            =   150
         TabIndex        =   14
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
            TabIndex        =   20
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
            TabIndex        =   18
            Top             =   195
            Width           =   1260
         End
         Begin MSMask.MaskEdBox TXTINVDATE 
            Height          =   315
            Left            =   4935
            TabIndex        =   19
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
            TabIndex        =   30
            Top             =   480
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   0
            TabIndex        =   29
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
            TabIndex        =   21
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
            TabIndex        =   17
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
            TabIndex        =   16
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
            TabIndex        =   15
            Top             =   180
            Width           =   480
         End
      End
      Begin VB.Frame FRMECONTROLS 
         BackColor       =   &H00C0E0FF&
         Height          =   2115
         Left            =   135
         TabIndex        =   10
         Top             =   6045
         Width           =   7755
         Begin VB.CommandButton CmdDeleteAll 
            Caption         =   "De&lete All"
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
            Left            =   6645
            TabIndex        =   8
            Top             =   1635
            Width           =   1050
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
            TabIndex        =   27
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
            TabIndex        =   23
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
            Left            =   60
            TabIndex        =   2
            Top             =   1635
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
            Left            =   2370
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
            Left            =   1200
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
            Left            =   3405
            TabIndex        =   5
            Top             =   1635
            Width           =   960
         End
         Begin MSDataListLib.DataList DataList2 
            Height          =   840
            Left            =   735
            TabIndex        =   24
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
            Caption         =   "TOTAL INCOME"
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
            Left            =   5970
            TabIndex        =   32
            Top             =   840
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
            Left            =   5940
            TabIndex        =   31
            Top             =   1080
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
            TabIndex        =   28
            Top             =   195
            Width           =   2040
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Income Head"
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
            TabIndex        =   25
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
            TabIndex        =   13
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
            TabIndex        =   12
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
            Left            =   4710
            TabIndex        =   11
            Top             =   1770
            Visible         =   0   'False
            Width           =   1080
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   5265
         Left            =   150
         TabIndex        =   26
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
Attribute VB_Name = "frmIncome"
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
    
    If grdsales.rows <= Val(TXTSLNO.text) Then grdsales.rows = grdsales.rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.text), 0) = Val(TXTSLNO.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 1) = DataList2.BoundText
    grdsales.TextMatrix(Val(TXTSLNO.text), 2) = Trim(DataList2.text)
    grdsales.TextMatrix(Val(TXTSLNO.text), 3) = Format(Val(TXTAMOUNT.text), "0.00")
    grdsales.TextMatrix(Val(TXTSLNO.text), 4) = Trim(TXTREMARKS.text)
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.rows - 1
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 3))
    Next i
    
    TXTSLNO.text = grdsales.rows
    TXTDEALER.text = ""
    TXTAMOUNT.text = ""
    TXTREMARKS.text = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    cmdRefresh.Enabled = True
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
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo + vbDefaultButton2, "INCOME ENTRY") = vbNo Then Exit Sub
    End If
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.text = "  /  /    "
    TXTSLNO.text = ""
    TXTDEALER.text = ""
    TXTAMOUNT.text = ""
    TXTREMARKS = ""
    LBLTOTAL.Caption = "0.00"
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
        'grdsales.TextMatrix(i, 0) = grdsales.TextMatrix(i + 1, 0)
        grdsales.TextMatrix(i, 1) = grdsales.TextMatrix(i + 1, 1)
        grdsales.TextMatrix(i, 2) = grdsales.TextMatrix(i + 1, 2)
        grdsales.TextMatrix(i, 3) = grdsales.TextMatrix(i + 1, 3)
        grdsales.TextMatrix(i, 4) = grdsales.TextMatrix(i + 1, 4)
    Next i
    grdsales.rows = grdsales.rows - 1
    TXTSLNO.text = Val(grdsales.rows)
    
    LBLTOTAL.Caption = ""
    For i = 1 To grdsales.rows - 1
        grdsales.TextMatrix(i, 0) = i
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(i, 3))
    Next i
    
    TXTDEALER.text = ""
    TXTAMOUNT.text = ""
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

Private Sub CmdDeleteAll_Click()
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE THE ENTIRE INCOME NO. " & txtBillNo.text, vbYesNo + vbDefaultButton2, "DELETE.....") = vbNo Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    db.Execute "delete From TRXINCOME WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXINCMAST WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'IN' AND INV_TRX_TYPE = 'IN'"
    M_ADD = False
    Call cmdcancel_Click
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description

End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDMODIFY_Click()
    
    If Val(TXTSLNO.text) >= grdsales.rows Then Exit Sub
    
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
            TXTSLNO.text = grdsales.rows
            TXTDEALER.text = ""
            TXTAMOUNT.text = ""
            TXTREMARKS.text = ""
        
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

Private Sub cmdRefresh_Click()
    Dim rstMaxNo As ADODB.Recordset
    
    If Not IsDate(TXTINVDATE.text) Then
        MsgBox "Enter Purchase Order Date", vbOKOnly, "EzBiz"
        Exit Sub
    End If
    On Error GoTo ErrHand
    Call appendpurchase
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From TRXINCMAST", db, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    
    grdsales.rows = 1
    TXTSLNO.text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.text = TXTLASTBILL.text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    lbldealer.Caption = ""
    flagchange.Caption = ""
    TXTDEALER.text = ""
    TXTINVDATE.text = "  /  /    "
    TXTSLNO.text = ""
    TXTAMOUNT.text = ""
    LBLTOTAL.Caption = "0.00"
    TXTREMARKS.text = ""
    grdsales.rows = 1
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
    Screen.MousePointer = vbNormal
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "INCOME ENTRY"
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
    
'    Set TRXMAST = New ADODB.Recordset
'    TRXMAST.Open "Select MAX(VCH_NO) From TRXINCOME", db, adOpenStatic, adLockReadOnly
'    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
'        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
'        TXTLASTBILL.Text = txtBillNo.Text
'    End If
'    TRXMAST.Close
'    Set TRXMAST = Nothing
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From TRXINCMAST", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.text = txtBillNo.text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
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
    grdsales.TextArray(2) = "EXPENSE HEAD"
    grdsales.TextArray(3) = "AMOUNT"
    grdsales.TextArray(4) = "REMARKS"

    PHYFLAG = True
    TXTDEALER.Enabled = False
    DataList2.Enabled = False
    TXTAMOUNT.Enabled = False
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
    Me.Width = 7935
    Me.Height = 9435
    Me.Left = 0
    Me.Top = 0
    LBLTOTAL.Caption = "0.00"
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
    End If
    Cancel = CLOSEALL
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    LBLTOTAL.Caption = ""
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.rows = 1
            i = 0
            grdsales.rows = 1
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRXINCOME WHERE VCH_NO = " & Val(txtBillNo.text) & " ORDER BY VCH_NO, VCH_DATE", db, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.rows = grdsales.rows + 1
                grdsales.FixedRows = 1
                i = i + 1
                grdsales.TextMatrix(i, 0) = i
                grdsales.TextMatrix(i, 1) = rstTRXMAST!ACT_CODE
                grdsales.TextMatrix(i, 2) = rstTRXMAST!ACT_NAME
                grdsales.TextMatrix(i, 3) = Format(rstTRXMAST!VCH_AMOUNT, "0.00")
                grdsales.TextMatrix(i, 4) = rstTRXMAST!REMARKS
                LBLTOTAL.Caption = Format(Val(LBLTOTAL.Caption) + rstTRXMAST!VCH_AMOUNT, "0.00")
                TXTINVDATE.text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
                rstTRXMAST.MoveNext
            Loop
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
                TXTINVDATE.SetFocus
            End If
    End Select
    
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    MsgBox err.Description
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
            'CMBDISTI.SetFocus
    End Select
End Sub

Private Sub TXTINVDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTAMOUNT_GotFocus()
    TXTAMOUNT.SelStart = 0
    TXTAMOUNT.SelLength = Len(TXTAMOUNT.text)
End Sub

Private Sub TXTAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTAMOUNT.text) = 0 Then Exit Sub
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTAMOUNT_LostFocus()
    TXTAMOUNT.text = Format(TXTAMOUNT.text, ".00")
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
                'TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTDEALER.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                DataList2.text = grdsales.TextMatrix(Val(TXTSLNO.text), 2)
                Call DataList2_Click
                TXTAMOUNT.text = Val(grdsales.TextMatrix(Val(TXTSLNO.text), 3))
                TXTREMARKS.text = Trim(grdsales.TextMatrix(Val(TXTSLNO.text), 4))
                
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
                TXTSLNO.text = Val(grdsales.rows)
                TXTDEALER.text = ""
                TXTAMOUNT.text = ""
                TXTREMARKS.text = ""
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

Public Sub appendpurchase()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rstMaxRec As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHand
    Screen.MousePointer = vbHourglass
    db.BeginTrans
    db.Execute "delete From TRXINCOME WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete From TRXINCMAST WHERE VCH_NO = " & Val(txtBillNo.text) & ""
    db.Execute "delete FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'IN' AND INV_TRX_TYPE = 'IN'"
    
    If grdsales.rows = 1 Then GoTo SKIP
    
    i = 0
    Set rstMaxRec = New ADODB.Recordset
    rstMaxRec.Open "Select MAX(REC_NO) From CASHATRXFILE ", db, adOpenForwardOnly
    If Not (rstMaxRec.EOF And rstMaxRec.BOF) Then
        i = IIf(IsNull(rstMaxRec.Fields(0)), 0, rstMaxRec.Fields(0))
    End If
    rstMaxRec.Close
    Set rstMaxRec = Nothing

    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM CASHATRXFILE WHERE INV_NO = " & Val(txtBillNo.text) & " AND INV_TYPE = 'IN' AND INV_TRX_TYPE = 'IN'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!REC_NO = i + 1
        RSTITEMMAST!INV_TYPE = "IN"
        RSTITEMMAST!INV_TRX_TYPE = "IN"
        RSTITEMMAST!INV_NO = Val(txtBillNo.text)
    End If
    RSTITEMMAST!TRX_TYPE = "CR"
    RSTITEMMAST!ACT_CODE = DataList2.BoundText
    If grdsales.rows > 1 Then RSTITEMMAST!ACT_NAME = Trim(grdsales.TextMatrix(grdsales.rows - 1, 2)) & " -" & Trim(grdsales.TextMatrix(grdsales.rows - 1, 4))
    RSTITEMMAST!AMOUNT = Val(LBLTOTAL.Caption)
    RSTITEMMAST!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTITEMMAST!ENTRY_DATE = Format(Date, "DD/MM/YYYY")
    RSTITEMMAST!check_flag = "S"
    RSTITEMMAST.Update
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from TRXINCMAST", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "IN"
    RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = ""
    RSTTRXFILE!ACT_NAME = ""
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
        RSTTRXFILE.Open "SELECT * from TRXINCOME", db, adOpenStatic, adLockOptimistic, adCmdText
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "IN"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.text, "DD/MM/YYYY")
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
    
ErrHand:
    Screen.MousePointer = vbNormal
    'If Err.Number <> -2147168237 Then
        MsgBox err.Description
    'End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='741')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from ACTMAST  WHERE (Mid(ACT_CODE, 1, 3)='741')And (LENGTH(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.text & "%'ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
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
ErrHand:
    MsgBox err.Description
    
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
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
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

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Expense head From List", vbOKOnly, "Expense Entry..."
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
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtremarks_GotFocus()
    TXTREMARKS.SelStart = 0
    TXTREMARKS.SelLength = Len(TXTREMARKS.text)
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
        Case Asc("'")
            KeyAscii = 0
'        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
'            KeyAscii = Asc(UCase(Chr(KeyAscii)))
'        Case Else
'            KeyAscii = 0
    End Select
End Sub

Private Sub TXTREMARKS_LostFocus()
    TXTREMARKS.text = Format(TXTREMARKS.text, ".00")
End Sub


