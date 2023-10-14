VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmExpenseStaffwo 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Expenses for Staff"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7845
   ControlBox      =   0   'False
   Icon            =   "frmExpenseStaffwo.frx":0000
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
      Left            =   1155
      TabIndex        =   21
      Top             =   270
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
      BackColor       =   &H00C0FFFF&
      Height          =   8235
      Left            =   -120
      TabIndex        =   8
      Top             =   -45
      Width           =   7920
      Begin MSFlexGridLib.MSFlexGrid grdsales 
         Height          =   4365
         Left            =   135
         TabIndex        =   25
         Top             =   1755
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   7699
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
      Begin VB.Frame FRMEMASTER 
         BackColor       =   &H00C0FFFF&
         Height          =   1605
         Left            =   135
         TabIndex        =   13
         Top             =   150
         Width           =   7740
         Begin VB.TextBox TXTEMPLOYEE 
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
            Left            =   1140
            TabIndex        =   32
            Top             =   525
            Width           =   3735
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
            Left            =   3120
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
         Begin MSDataListLib.DataList Dlstemployee 
            Height          =   645
            Left            =   1140
            TabIndex        =   33
            Top             =   870
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
         Begin VB.Label empflag 
            Height          =   315
            Left            =   5685
            TabIndex        =   36
            Top             =   870
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.Label lblemployee 
            Height          =   315
            Left            =   5415
            TabIndex        =   35
            Top             =   1170
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "EMPLOYEE"
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
            Left            =   105
            TabIndex        =   34
            Top             =   585
            Width           =   1005
         End
         Begin VB.Label lbldealer 
            Height          =   315
            Left            =   6105
            TabIndex        =   29
            Top             =   570
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label flagchange 
            Height          =   315
            Left            =   7095
            TabIndex        =   28
            Top             =   930
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
            Left            =   2055
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
            Left            =   105
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
         BackColor       =   &H00C0FFFF&
         Height          =   2115
         Left            =   135
         TabIndex        =   9
         Top             =   6045
         Width           =   7755
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
            Left            =   5715
            MaxLength       =   30
            TabIndex        =   26
            Top             =   450
            Width           =   1980
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
            Left            =   4425
            MaxLength       =   8
            TabIndex        =   1
            Top             =   450
            Width           =   1275
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
            Caption         =   "TOTAL EXPENSE"
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
            Left            =   5910
            TabIndex        =   31
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
            Left            =   5700
            TabIndex        =   30
            Top             =   1080
            Width           =   1995
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
            Left            =   5715
            TabIndex        =   27
            Top             =   195
            Width           =   1980
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            Caption         =   "Expense Head"
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
            Left            =   4425
            TabIndex        =   11
            Top             =   195
            Width           =   1275
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
   End
End
Attribute VB_Name = "frmExpenseStaffwo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim EMP_REC As New ADODB.Recordset
Dim EMP_FLAG As Boolean
Dim PHYFLAG As Boolean
Dim ACT_FLAG As Boolean
Dim CLOSEALL As Integer
Dim M_EDIT As Boolean
Dim M_ADD As Boolean

Private Sub CMDADD_Click()
    Dim I As Integer
    
    If grdsales.Rows <= Val(TXTSLNO.Text) Then grdsales.Rows = grdsales.Rows + 1
    grdsales.FixedRows = 1
    grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = Val(TXTSLNO.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = DataList2.BoundText
    grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = Trim(DataList2.Text)
    grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = Format(Val(TXTAMOUNT.Text), "0.00")
    grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = Trim(txtremarks.Text)
    
    LBLTOTAL.Caption = ""
    For I = 1 To grdsales.Rows - 1
        LBLTOTAL.Caption = Val(LBLTOTAL.Caption) + Val(grdsales.TextMatrix(I, 3))
    Next I
    
    TXTSLNO.Text = grdsales.Rows
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    txtremarks.Text = ""
    cmdadd.Enabled = False
    CmdDelete.Enabled = False
    CMDEXIT.Enabled = False
    M_EDIT = False
    M_ADD = True
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    txtBillNo.Enabled = False
    
    If grdsales.Rows >= 18 Then grdsales.TopRow = grdsales.Rows - 1
        
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
        If MsgBox("Changes have been made. Do you want to Cancel?", vbYesNo, "PURCHASE ORDER...") = vbNo Then Exit Sub
    End If
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    TXTINVDATE.Text = Date
    TXTSLNO.Text = ""
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    txtremarks.Text = ""
    LBLTOTAL.Caption = "0.00"
    grdsales.Rows = 1
    empflag.Caption = ""
    TXTEMPLOYEE.Text = ""
    lblemployee.Caption = ""
    CMDEXIT.Enabled = True
    txtBillNo.SetFocus
    M_ADD = False
End Sub

Private Sub CmdDelete_Click()
    Dim I As Integer
    
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
    
    For I = Val(TXTSLNO.Text) - 1 To grdsales.Rows - 2
        grdsales.TextMatrix(Val(TXTSLNO.Text), 0) = I
        grdsales.TextMatrix(Val(TXTSLNO.Text), 1) = grdsales.TextMatrix(I + 1, 1)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 2) = grdsales.TextMatrix(I + 1, 2)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 3) = grdsales.TextMatrix(I + 1, 3)
        grdsales.TextMatrix(Val(TXTSLNO.Text), 4) = grdsales.TextMatrix(I + 1, 4)
    Next I
    grdsales.Rows = grdsales.Rows - 1
    TXTSLNO.Text = Val(grdsales.Rows)
    TXTDEALER.Text = ""
    TXTAMOUNT.Text = ""
    txtremarks = ""
    cmdadd.Enabled = False
    TXTSLNO.Enabled = True
    TXTSLNO.SetFocus
    CmdDelete.Enabled = False
    CMDMODIFY.Enabled = False
    CMDEXIT.Enabled = False
    M_ADD = True
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
    TXTAMOUNT.Enabled = True
    TXTAMOUNT.SetFocus
    
End Sub

Private Sub CMDMODIFY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TXTSLNO.Text = grdsales.Rows
            TXTDEALER.Text = ""
            TXTAMOUNT.Text = ""
            txtremarks.Text = ""
        
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            TXTDEALER.Enabled = False
            DataList2.Enabled = False
            TXTAMOUNT.Enabled = False
            txtremarks.Enabled = False
            CMDMODIFY.Enabled = False
            CmdDelete.Enabled = False
            M_EDIT = False
    End Select
End Sub

Private Sub cmdRefresh_Click()
    Dim rstMaxNo As ADODB.Recordset
    
    If Not IsDate(TXTINVDATE.Text) Then
        MsgBox "Please Enter a Valid Date", vbOKOnly, "Expense Entry..."
        TXTINVDATE.SetFocus
        Exit Sub
    End If
    
    If Dlstemployee.Text = "" Then
        MsgBox "Select Employee From List", vbOKOnly, "Expense Entry..."
        TXTEMPLOYEE.SetFocus
        Exit Sub
    End If
    
    If IsNull(Dlstemployee.SelectedItem) Then
        MsgBox "Select Employee From List", vbOKOnly, "Expense Entry..."
        TXTEMPLOYEE.SetFocus
        Exit Sub
    End If
    On Error GoTo eRRhAND
    Call appendpurchase
    
    Set rstMaxNo = New ADODB.Recordset
    rstMaxNo.Open "Select MAX(VCH_NO) From TRXEXP_MAST", db2, adOpenStatic, adLockReadOnly
    If Not (rstMaxNo.EOF And rstMaxNo.BOF) Then
        txtBillNo.Text = IIf(IsNull(rstMaxNo.Fields(0)), 1, rstMaxNo.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    rstMaxNo.Close
    Set rstMaxNo = Nothing
    grdsales.Rows = 1
    TXTSLNO.Text = 1
    cmdRefresh.Enabled = False
    txtBillNo.Enabled = True
    txtBillNo.Text = TXTLASTBILL.Text
    FRMEMASTER.Enabled = False
    FRMECONTROLS.Enabled = False
    lbldealer.Caption = ""
    flagchange.Caption = ""
    lblemployee.Caption = ""
    empflag.Caption = ""
    TXTDEALER.Text = ""
    TXTINVDATE.Text = Date
    TXTSLNO.Text = ""
    TXTAMOUNT.Text = ""
    LBLTOTAL.Caption = "0.00"
    txtremarks.Text = ""
    grdsales.Rows = 1
    CMDEXIT.Enabled = True
    empflag.Caption = ""
    TXTEMPLOYEE.Text = ""
    lblemployee.Caption = ""
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
    txtBillNo.SetFocus
    M_ADD = False
    MsgBox "SAVED SUCCESSFULLY", vbOKOnly, "PURCHASE ORDER"
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub Form_Activate()
    On Error GoTo eRRhAND
    txtBillNo.SetFocus
    Exit Sub
eRRhAND:
    If Err.Number = 5 Then Exit Sub
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(Val(VCH_NO)) From TRXEXP_MAST", db2, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        txtBillNo.Text = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
        TXTLASTBILL.Text = txtBillNo.Text
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    ACT_FLAG = True
    EMP_FLAG = True
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
    TMPFLAG = True
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
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If PHYFLAG = False Then PHY.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If EMP_FLAG = False Then EMP_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub TXTBILLNO_GotFocus()
    txtBillNo.SelStart = 0
    txtBillNo.SelLength = Len(txtBillNo.Text)
End Sub

Private Sub TXTBILLNO_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rstTRXMAST As ADODB.Recordset
    Dim RSTDIST As ADODB.Recordset
    Dim RSTTRNSMAST As ADODB.Recordset
    Dim I As Integer

    On Error GoTo eRRhAND
    LBLTOTAL.Caption = ""
    Select Case KeyCode
        Case vbKeyReturn
            grdsales.Rows = 1
            I = 0
                        
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRXEXP_MAST WHERE VCH_NO = " & Val(txtBillNo.Text) & "", db2, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF Or rstTRXMAST.BOF) Then
                LBLTOTAL.Caption = Format(rstTRXMAST!VCH_AMOUNT, "0.00")
                TXTEMPLOYEE.Text = rstTRXMAST!ACT_NAME
                TXTINVDATE.Text = Format(rstTRXMAST!VCH_DATE, "DD/MM/YYYY")
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "Select * From TRXFILE_EXP WHERE VCH_NO = " & Val(txtBillNo.Text) & " ORDER BY [VCH_NO], [VCH_DATE]", db2, adOpenStatic, adLockReadOnly
            Do Until rstTRXMAST.EOF
                grdsales.Rows = grdsales.Rows + 1
                grdsales.FixedRows = 1
                I = I + 1
                grdsales.TextMatrix(I, 0) = I
                grdsales.TextMatrix(I, 1) = rstTRXMAST!EXP_CODE
                grdsales.TextMatrix(I, 2) = rstTRXMAST!EXP_NAME
                grdsales.TextMatrix(I, 3) = Format(rstTRXMAST!TRX_TOTAL, "0.00")
                grdsales.TextMatrix(I, 4) = rstTRXMAST!REMARKS
                rstTRXMAST.MoveNext
            Loop
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            
            TXTSLNO.Text = grdsales.Rows
            TXTSLNO.Enabled = True
            txtBillNo.Enabled = False
            FRMEMASTER.Enabled = True
            If I > 0 Or (Val(txtBillNo.Text) < Val(TXTLASTBILL.Text)) Then
                'FRMEMASTER.Enabled = False
                cmdcancel.SetFocus
            Else
                TXTINVDATE.SetFocus
            End If
    End Select
    Dlstemployee.Text = TXTEMPLOYEE.Text
    Call Dlstemployee_Click
    Exit Sub
eRRhAND:
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
    If Val(txtBillNo.Text) = 0 Or Val(txtBillNo.Text) > Val(TXTLASTBILL.Text) Then txtBillNo.Text = TXTLASTBILL.Text
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
                TXTEMPLOYEE.SetFocus
'                FRMECONTROLS.Enabled = True
'                TXTSLNO.Enabled = True
'                TXTSLNO.SetFocus
            End If
        Case vbKeyEscape
            If M_ADD = False Then
                FRMECONTROLS.Enabled = False
                FRMEMASTER.Enabled = False
                txtBillNo.Enabled = True
                txtBillNo.SetFocus
            End If
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
    TXTAMOUNT.SelLength = Len(TXTAMOUNT.Text)
End Sub

Private Sub TXTAMOUNT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTAMOUNT.Text) = 0 Then Exit Sub
            txtremarks.Enabled = True
            TXTAMOUNT.Enabled = False
            txtremarks.SetFocus
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
                'TXTPRODUCT.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 1)
                TXTDEALER.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                DataList2.Text = grdsales.TextMatrix(Val(TXTSLNO.Text), 2)
                Call DataList2_Click
                TXTAMOUNT.Text = Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 3))
                txtremarks.Text = Trim(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
                
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
            If CmdDelete.Enabled = True Then
                TXTSLNO.Text = Val(grdsales.Rows)
                TXTDEALER.Text = ""
                TXTAMOUNT.Text = ""
                txtremarks.Text = ""
                cmdadd.Enabled = False
                CmdDelete.Enabled = False
                TXTSLNO.Enabled = True
                TXTSLNO.SetFocus
            ElseIf grdsales.Rows > 1 Then
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
    Dim I As Integer
    
    On Error GoTo eRRhAND
    Screen.MousePointer = vbHourglass
    db2.Execute "delete * From TRXEXP_MAST WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    db2.Execute "delete * From TRXFILE_EXP WHERE VCH_NO = " & Val(txtBillNo.Text) & ""
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from [TRXEXP_MAST]", db2, adOpenStatic, adLockOptimistic, adCmdText
    RSTTRXFILE.AddNew
    RSTTRXFILE!TRX_TYPE = "EX"
    RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
    RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
    RSTTRXFILE!ACT_CODE = Dlstemployee.BoundText
    RSTTRXFILE!ACT_NAME = Trim(Dlstemployee.Text)
    RSTTRXFILE!VCH_AMOUNT = Val(LBLTOTAL.Caption)
    RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
    RSTTRXFILE!C_USER_ID = "SM"
    RSTTRXFILE!MODIFY_DATE = Format(Date, "dd/mm/yyyy")
    RSTTRXFILE!M_USER_ID = "SM"
    RSTTRXFILE.Update
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT * from [TRXFILE_EXP]", db2, adOpenStatic, adLockOptimistic, adCmdText
    For I = 1 To grdsales.Rows - 1
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "EX"
        RSTTRXFILE!VCH_NO = Val(txtBillNo.Text)
        RSTTRXFILE!VCH_DATE = Format(TXTINVDATE.Text, "DD/MM/YYYY")
        RSTTRXFILE!LINE_NO = I
        RSTTRXFILE!EXP_CODE = Trim(grdsales.TextMatrix(I, 1))
        RSTTRXFILE!EXP_NAME = Trim(grdsales.TextMatrix(I, 2))
        RSTTRXFILE!TRX_TOTAL = Val(grdsales.TextMatrix(I, 3))
        RSTTRXFILE!REMARKS = Trim(grdsales.TextMatrix(I, 4))
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!MODIFY_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!M_USER_ID = "SM"
        RSTTRXFILE.Update
    Next I
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='641')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='641')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTDEALER.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
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
eRRhAND:
    MsgBox Err.Description
    
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
    txtremarks.SelStart = 0
    txtremarks.SelLength = Len(txtremarks.Text)
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txtremarks.Enabled = False
            cmdadd.Enabled = True
            cmdadd.SetFocus
        Case vbKeyEscape
            TXTAMOUNT.Enabled = True
            txtremarks.Enabled = False
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
    txtremarks.Text = Format(txtremarks.Text, ".00")
End Sub

Private Sub TXTEMPLOYEE_Change()
    On Error GoTo eRRhAND
    If empflag.Caption <> "1" Then
        If EMP_FLAG = True Then
            EMP_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='321')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTEMPLOYEE.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        Else
            EMP_REC.Close
            EMP_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='321')And (len(ACT_CODE)>3) And ACT_NAME Like '" & Me.TXTEMPLOYEE.Text & "%'ORDER BY ACT_NAME", db2, adOpenStatic, adLockReadOnly, adCmdText
            EMP_FLAG = False
        End If
        If (EMP_REC.EOF And EMP_REC.BOF) Then
            lblemployee.Caption = ""
        Else
            lblemployee.Caption = EMP_REC!ACT_NAME
        End If
        Set Me.Dlstemployee.RowSource = EMP_REC
        Dlstemployee.ListField = "ACT_NAME"
        Dlstemployee.BoundColumn = "ACT_CODE"
    End If
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub TXTEMPLOYEE_GotFocus()
    TXTEMPLOYEE.SelStart = 0
    TXTEMPLOYEE.SelLength = Len(TXTEMPLOYEE.Text)
End Sub

Private Sub TXTEMPLOYEE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Dlstemployee.VisibleCount = 0 Then Exit Sub
            Dlstemployee.SetFocus
        Case vbKeyEscape
            TXTINVDATE.SetFocus
    End Select

End Sub

Private Sub TXTEMPLOYEE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Dlstemployee_GotFocus()
    empflag.Caption = 1
    TXTEMPLOYEE = lblemployee.Caption
    Dlstemployee.Text = TXTEMPLOYEE.Text
    Call Dlstemployee_Click
End Sub

Private Sub Dlstemployee_LostFocus()
     empflag.Caption = ""
End Sub

Private Sub Dlstemployee_Click()
    TXTEMPLOYEE = Dlstemployee.Text
    lblemployee.Caption = TXTEMPLOYEE
End Sub

Private Sub Dlstemployee_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Dlstemployee.Text = "" Then Exit Sub
            If IsNull(Dlstemployee.SelectedItem) Then
                MsgBox "Select Expense head From List", vbOKOnly, "Expense Entry..."
                Dlstemployee.SetFocus
                Exit Sub
            End If
            FRMECONTROLS.Enabled = True
            TXTSLNO.Enabled = True
            TXTSLNO.SetFocus
            'FRMEHEAD.Enabled = False
            'TXTSLNO.Enabled = True
            'TXTSLNO.SetFocus
        Case vbKeyEscape
            TXTEMPLOYEE.SetFocus
    End Select
End Sub

Private Sub Dlstemployee_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
