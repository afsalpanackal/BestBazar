VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAstmaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assets Item Creation"
   ClientHeight    =   5400
   ClientLeft      =   -15
   ClientTop       =   270
   ClientWidth     =   14430
   ControlBox      =   0   'False
   Icon            =   "frmAstmast.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   14430
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1320
      Top             =   9285
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Height          =   300
      Left            =   8520
      ScaleHeight     =   300
      ScaleWidth      =   750
      TabIndex        =   24
      Top             =   120
      Width           =   750
   End
   Begin VB.TextBox TxtProduct 
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
      Height          =   435
      Left            =   1320
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   4965
   End
   Begin VB.TextBox TxtItemcode 
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
      Height          =   435
      Left            =   1320
      MaxLength       =   21
      TabIndex        =   16
      Top             =   360
      Width           =   4950
   End
   Begin VB.Frame FRAME 
      Height          =   4650
      Left            =   15
      TabIndex        =   1
      Top             =   750
      Width           =   14430
      Begin VB.ComboBox CmbPack 
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
         ItemData        =   "frmAstmast.frx":000C
         Left            =   1050
         List            =   "frmAstmast.frx":0034
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   3510
         Width           =   1095
      End
      Begin VB.CommandButton CmdSave 
         BackColor       =   &H00400000&
         Caption         =   "&SAVE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   3960
         MaskColor       =   &H80000007&
         TabIndex        =   29
         Top             =   3375
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H00400000&
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
         Height          =   540
         Left            =   6630
         MaskColor       =   &H80000007&
         TabIndex        =   11
         Top             =   3375
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00400000&
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
         Height          =   540
         Left            =   5295
         MaskColor       =   &H80000007&
         TabIndex        =   10
         Top             =   3375
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox txtcategory 
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
         ForeColor       =   &H00004080&
         Height          =   390
         Left            =   1050
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1200
         Width           =   2895
      End
      Begin VB.CheckBox chknewcategory 
         Caption         =   "N&ew Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   1050
         TabIndex        =   8
         Top             =   3075
         Width           =   1725
      End
      Begin VB.CommandButton CMDDELETE 
         BackColor       =   &H00400000&
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6630
         MaskColor       =   &H80000007&
         TabIndex        =   13
         Top             =   4005
         UseMaskColor    =   -1  'True
         Width           =   1275
      End
      Begin VB.TextBox TXTITEM 
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
         Height          =   450
         Left            =   2040
         TabIndex        =   2
         Top             =   435
         Width           =   5910
      End
      Begin MSDataListLib.DataList Datacategory 
         Height          =   1320
         Left            =   1050
         TabIndex        =   12
         Top             =   1605
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   2328
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   16512
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
      Begin VB.Frame FrmeCompany 
         BorderStyle     =   0  'None
         Height          =   2520
         Left            =   4020
         TabIndex        =   3
         Top             =   1050
         Width           =   3945
         Begin VB.CheckBox chknewcomp 
            Caption         =   "&New Company"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   1035
            TabIndex        =   5
            Top             =   1995
            Width           =   1695
         End
         Begin VB.TextBox txtcompany 
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
            ForeColor       =   &H00004080&
            Height          =   405
            Left            =   1035
            MaxLength       =   25
            TabIndex        =   4
            Top             =   105
            Width           =   2895
         End
         Begin MSDataListLib.DataList Datacompany 
            Height          =   1320
            Left            =   1035
            TabIndex        =   6
            Top             =   525
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   2328
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16512
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
         Begin VB.Label LBLLP 
            Height          =   375
            Left            =   2760
            TabIndex        =   25
            Top             =   2280
            Width           =   1215
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Company"
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
            Left            =   0
            TabIndex        =   7
            Top             =   210
            Width           =   960
         End
      End
      Begin MSDataGridLib.DataGrid grdtmp 
         Height          =   4425
         Left            =   7965
         TabIndex        =   26
         Top             =   135
         Width           =   6420
         _ExtentX        =   11324
         _ExtentY        =   7805
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
            Weight          =   400
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   3570
         Width           =   825
      End
      Begin VB.Label lblPack 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   2130
         TabIndex        =   28
         Top             =   5055
         Width           =   465
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
         Index           =   2
         Left            =   45
         TabIndex        =   15
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Description"
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
         Index           =   1
         Left            =   75
         TabIndex        =   14
         Top             =   495
         Width           =   1995
      End
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   1320
      TabIndex        =   17
      Top             =   810
      Visible         =   0   'False
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   2858
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
   Begin VB.Label LBLITEMNAME 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   480
      Left            =   8910
      TabIndex        =   27
      Top             =   120
      Width           =   5205
   End
   Begin VB.Label lbldealer 
      Height          =   315
      Left            =   6600
      TabIndex        =   23
      Top             =   375
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label flagchange 
      Height          =   315
      Left            =   7680
      TabIndex        =   22
      Top             =   210
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   7290
      TabIndex        =   21
      Top             =   60
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   6420
      TabIndex        =   20
      Top             =   60
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ITEM CODE"
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
      Left            =   195
      TabIndex        =   19
      Top             =   435
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Press F3 to Search...... Press Esc to Exit.."
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
      Left            =   1320
      TabIndex        =   18
      Top             =   0
      Width           =   3960
   End
End
Attribute VB_Name = "frmAstmaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY As New ADODB.Recordset
Dim PHYFLAG As Boolean
Dim REPFLAG As Boolean
Dim COMPANYFLAG As Boolean
Dim CATEGORYFLAG As Boolean
Dim RSTREP As New ADODB.Recordset
Dim RSTCOMPANY As New ADODB.Recordset
Dim RSTCATEGORY As New ADODB.Recordset

Private Sub chknewcategory_Click()
    On Error Resume Next
    txtcategory.SetFocus
End Sub

Private Sub chknewcomp_Click()
    On Error Resume Next
    txtcompany.SetFocus
End Sub

Private Sub CmbPack_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CmbPack.ListIndex = -1 Then CmbPack.ListIndex = 0
            CmdSave.SetFocus
        Case vbKeyEscape
            If FrmeCompany.Visible = True Then
                txtcompany.SetFocus
            Else
                txtcategory.SetFocus
            End If
    End Select
End Sub

Private Sub CmbPack_LostFocus()
    LblPack.Caption = CmbPack.Text
End Sub

Private Sub cmdcancel_Click()
    
    TXTPRODUCT.Text = ""
    TXTITEM.Text = ""
    lblitemname.Caption = ""
    txtcategory.Text = ""
    txtcompany.Text = ""
    CmbPack.ListIndex = -1
    TXTITEMCODE.Enabled = True
    DataList2.Enabled = True
    Frame.Visible = False
    TXTPRODUCT.Visible = False
    DataList2.Visible = False
    TXTITEMCODE.SetFocus
    chknewcategory.value = 0
    chknewcomp.value = 0
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ASTMAST ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim RSTITEMMAST As ADODB.Recordset
    On Error Resume Next
    If TXTITEM.Text = "" Then
        MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
        TXTITEM.SetFocus
        Exit Sub
    End If
    If txtcompany.Visible = True Then
        If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And Trim(txtcompany.Text) = "" Then
            MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
            txtcompany.SetFocus
            Exit Sub
        End If
        If UCase(Datacategory.BoundText) <> "SERVICE CHARGE" And chknewcomp.value = 0 And Datacompany.BoundText = "" Then
            MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
            txtcompany.SetFocus
            Exit Sub
        End If
    End If
    If Trim(txtcategory.Text) = "" Then
        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
        txtcategory.SetFocus
        Exit Sub
    End If
    
    If chknewcategory.value = 0 And Datacategory.BoundText = "" Then
        MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
        txtcategory.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRhAND
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ASTMAST WHERE ITEM_NAME = '" & Trim(TXTITEM.Text) & "' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        MsgBox "The Item name already exists...", vbOKOnly, "Item Master"
        RSTITEMMAST.Close
        Set RSTITEMMAST = Nothing
        Exit Sub
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ASTMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.value = 1 Then RSTITEMMAST!Category = txtcategory.Text Else RSTITEMMAST!Category = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Datacompany.BoundText
        RSTITEMMAST!PACK_TYPE = CmbPack.Text
        RSTITEMMAST!SALES_TAX = 0 'Val(TxtTax.Text)
        RSTITEMMAST!REMARKS = "" 'Trim(TxtHSN.Text)
        RSTITEMMAST!CHECK_FLAG = "V"
        RSTITEMMAST.Update
    Else
        RSTITEMMAST.AddNew
        RSTITEMMAST!ITEM_CODE = Trim(TXTITEMCODE.Text)
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        If chknewcategory.value = 1 Then RSTITEMMAST!Category = txtcategory.Text Else RSTITEMMAST!Category = Datacategory.BoundText
        RSTITEMMAST!UNIT = 1
        If chknewcomp.value = 1 Then RSTITEMMAST!MANUFACTURER = txtcompany.Text Else RSTITEMMAST!MANUFACTURER = Trim(Datacompany.BoundText)
        RSTITEMMAST!REMARKS = "" 'Trim(TxtHSN.Text)
        RSTITEMMAST!PACK_TYPE = CmbPack.Text
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
        RSTITEMMAST!CHECK_FLAG = "V"
        RSTITEMMAST!ITEM_COST = 0
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT MANUFACTURER FROM MANUFACT WHERE MANUFACTURER = '" & Trim(txtcompany.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!MANUFACTURER = Trim(txtcompany.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT CATEGORY FROM CATEGORY WHERE CATEGORY = '" & Trim(txtcategory.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST.AddNew
        RSTITEMMAST!Category = Trim(txtcategory.Text)
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ASTRXFILE WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        RSTITEMMAST!ITEM_NAME = Trim(TXTITEM.Text)
        RSTITEMMAST!MFGR = Trim(txtcompany.Text)
        RSTITEMMAST!Category = Trim(txtcategory.Text)
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MsgBox "SAVED SUCCESSFULLY..", vbOKOnly, "ITEM CREATION"
    cmdcancel_Click
Exit Sub
eRRhAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub CmdSave_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRhAND
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT ITEM_CODE FROM ASTMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                TXTITEMCODE.Text = RSTITEMMAST!ITEM_CODE
            End If
            TXTPRODUCT.Visible = False
            DataList2.Visible = False
            TXTITEMCODE.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.SetFocus
    End Select
    Exit Sub
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Call txtcompany_Change
    TXTITEMCODE.SetFocus
End Sub

Private Sub Form_Load()
    
    Dim TRXMAST As ADODB.Recordset
    On Error GoTo eRRhAND
    
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(CONVERT(ITEM_CODE, SIGNED INTEGER)) From ASTMAST ", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        If IsNull(TRXMAST.Fields(0)) Then
            TXTITEMCODE.Text = 1
        Else
            TXTITEMCODE.Text = Val(TRXMAST.Fields(0)) + 1
        End If
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing

    PHYFLAG = True
    REPFLAG = True
    COMPANYFLAG = True
    CATEGORYFLAG = True
    Call txtcategory_Change
    'TMPFLAG = True
    'Width = 8385
    'Height = 4575
    Left = 3500
    Top = 0
    Frame.Visible = False
    'txtunit.Visible = False
    
    Picture2.ScaleMode = 3
    Picture2.Height = Picture2.Height * (1.4 * 40 / Picture2.ScaleHeight)
    Picture2.FontSize = 8

    Exit Sub
eRRhAND:
    MsgBox (Err.Description)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If REPFLAG = False Then RSTREP.Close
    If COMPANYFLAG = False Then RSTCOMPANY.Close
    If CATEGORYFLAG = False Then RSTCATEGORY.Close
    If PHYFLAG = False Then PHY.Close
    'If TMPFLAG = False Then rstTMP.Close
    'MDIMAIN.Enabled = True
End Sub

Private Sub grdtmp_DblClick()
    On Error Resume Next
    TXTITEM.Text = Trim(grdtmp.Columns(1))
    TXTITEM.SetFocus
End Sub

Private Sub grdtmp_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            On Error Resume Next
            TXTITEM.Text = Trim(grdtmp.Columns(1))
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub TXTITEM_Change()
    On Error GoTo eRRhAND
    Set grdtmp.DataSource = Nothing
    If PHYFLAG = True Then
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ASTMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    Else
        PHY.Close
        'PHY.Open "Select ITEM_NAME, CATEGORY FROM ASTMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ASTMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
        PHYFLAG = False
    End If
    Set grdtmp.DataSource = PHY
    grdtmp.Columns(0).Caption = "Code"
    'grdtmp.Columns(8).Caption = ""
    
    grdtmp.Columns(0).Width = 1000
    grdtmp.Columns(1).Width = 3800
    grdtmp.Columns(2).Width = 1200
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTITEM.Text = "" Then
                MsgBox "ENTER NAME OF PRODUCT", vbOKOnly, "PRODUCT MASTER"
                TXTITEM.SetFocus
                Exit Sub
            End If
            txtcategory.SetFocus
    End Select
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTITEMCODE_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
End Sub

Private Sub TXTITEMCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTITEMMAST As ADODB.Recordset
    Select Case KeyCode
        Case vbKeyReturn
            On Error GoTo eRRhAND
            If Trim(TXTITEMCODE.Text) = "" Then Exit Sub
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * FROM ASTMAST WHERE ITEM_CODE = '" & Trim(TXTITEMCODE.Text) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                On Error Resume Next
                
                TXTITEM.Text = IIf(IsNull(RSTITEMMAST!ITEM_NAME), "", RSTITEMMAST!ITEM_NAME)
                txtcategory.Text = IIf(IsNull(RSTITEMMAST!Category), "", RSTITEMMAST!Category)
                txtcompany.Text = IIf(IsNull(RSTITEMMAST!MANUFACTURER), "", RSTITEMMAST!MANUFACTURER)
                On Error Resume Next
                CmbPack.Text = IIf(IsNull(RSTITEMMAST!PACK_TYPE), 0, RSTITEMMAST!PACK_TYPE)
                On Error GoTo eRRhAND
                Datacategory.Text = txtcategory.Text
                Call Datacategory_Click
                Datacompany.Text = txtcompany.Text
                Call Datacompany_Click
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        
            TXTITEMCODE.Enabled = False
            Frame.Visible = True
            TXTITEM.SetFocus
        Case 114
            TXTPRODUCT.Visible = True
            DataList2.Visible = True
            TXTPRODUCT.SetFocus
        Case vbKeyEscape
            Call cmdexit_Click
    End Select
Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTITEMCODE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTPRODUCT_Change()
    On Error GoTo eRRhAND
    If REPFLAG = True Then
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ASTMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    Else
        RSTREP.Close
        RSTREP.Open "Select DISTINCT ITEM_CODE,ITEM_NAME From ASTMAST  WHERE ITEM_NAME Like '" & TXTPRODUCT.Text & "%'ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
        REPFLAG = False
    End If
    Set DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub TXTPRODUCT_GotFocus()
    TXTPRODUCT.SelStart = 0
    TXTPRODUCT.SelLength = Len(TXTPRODUCT.Text)
End Sub

Private Sub TXTPRODUCT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If TXTPRODUCT.Text = "" Then Exit Sub
            DataList2.SetFocus
        Case vbKeyEscape
            TXTPRODUCT.Visible = False
            DataList2.Visible = False
            TXTITEMCODE.SetFocus
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

Private Sub CmdDelete_Click()
    If MDIMAIN.StatusBar.Panels(9).Text = "Y" Then Exit Sub
    Dim rststock As ADODB.Recordset
    Dim i As Long
    
    i = 0
    If TXTITEMCODE.Text = "" Then Exit Sub
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from ASTRXFILE where ASTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        i = i + 1
    End If
    rststock.Close
    Set rststock = Nothing
    
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT BAL_QTY from ASTRXFILE where ASTRXFILE.ITEM_CODE = '" & TxtItemcode.Text & "'", db, adOpenForwardOnly
'    Do Until rststock.EOF
'        i = i + rststock!BAL_QTY
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
    
    If i <> 0 Then
        MsgBox "Cannot Delete " & TXTITEM.Text & " Since Transactions is Available", vbCritical, "DELETING ITEM...."
        Exit Sub
    End If
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & TXTITEM.Text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    db.Execute ("DELETE from ASTRXFILE where ASTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    db.Execute ("DELETE from ASTMAST where ASTMAST.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
    
    'tXTMEDICINE.Tag = tXTMEDICINE.Text
    'tXTMEDICINE.Text = ""
    'tXTMEDICINE.Text = tXTMEDICINE.Tag
    'TXTQTY.Text = ""
    MsgBox "ITEM " & TXTITEM.Text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    Call cmdcancel_Click
    Exit Sub
   
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub txtcategory_Change()

    On Error GoTo eRRhAND
    If flagchange.Caption <> "1" Then
        If CATEGORYFLAG = True Then
            RSTCATEGORY.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            CATEGORYFLAG = False
        Else
            RSTCATEGORY.Close
            RSTCATEGORY.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & txtcategory.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
            CATEGORYFLAG = False
        End If
        If (RSTCATEGORY.EOF And RSTCATEGORY.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = RSTCATEGORY!Category
        End If
        Set Me.Datacategory.RowSource = RSTCATEGORY
        Datacategory.ListField = "CATEGORY"
        Datacategory.BoundColumn = "CATEGORY"
    End If
    Exit Sub
eRRhAND:
    MsgBox Err.Description

End Sub

Private Sub txtcategory_GotFocus()
    txtcategory.SelStart = 0
    txtcategory.SelLength = Len(txtcategory.Text)
End Sub

Private Sub txtcategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If chknewcategory.value = 0 And Datacategory.BoundText = "" Then
                If Datacategory.VisibleCount = 0 Then Exit Sub
'                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'                Txtcompany.SetFocus
                Datacategory.SetFocus
                Exit Sub
            Else
                If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
                    FrmeCompany.Visible = False
                    CmbPack.Visible = False
                    
                    CmdSave.SetFocus
                Else
                    FrmeCompany.Visible = True
                    CmbPack.Visible = True
                    
                    txtcompany.SetFocus
                End If
            End If
        Case vbKeyEscape
            TXTITEM.SetFocus
    End Select
End Sub

Private Sub txtcategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacategory_Click()

    txtcategory.Text = Datacategory.Text
    lbldealer.Caption = txtcategory.Text

    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
        FrmeCompany.Visible = False
        CmbPack.Visible = False
        
    Else
        FrmeCompany.Visible = True
        CmbPack.Visible = True
        
    End If

End Sub

Private Sub Datacategory_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtcategory.Text = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If

            If chknewcategory.value = 0 And Datacategory.BoundText = "" Then
                MsgBox "SELECT CATEGORY", vbOKOnly, "PRODUCT MASTER"
                txtcategory.SetFocus
                Exit Sub
            End If
            If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
                FrmeCompany.Visible = False
                
                CmbPack.Visible = False
                
                CmdSave.SetFocus
            Else
                FrmeCompany.Visible = True
                CmbPack.Visible = True
                
                txtcompany.SetFocus
            End If
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select
End Sub

Private Sub Datacategory_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacategory_GotFocus()
    flagchange.Caption = 1
    txtcategory.Text = lbldealer.Caption
    Datacategory.Text = txtcategory.Text
    Call Datacategory_Click
End Sub

Private Sub Datacategory_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub txtcompany_Change()

    On Error GoTo eRRhAND
    If FLAGCHANGE2.Caption <> "1" Then
        If COMPANYFLAG = True Then
            RSTCOMPANY.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            COMPANYFLAG = False
        Else
            RSTCOMPANY.Close
            RSTCOMPANY.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & txtcompany.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
            COMPANYFLAG = False
        End If
        If (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
            LBLDEALER2.Caption = ""
        Else
            LBLDEALER2.Caption = RSTCOMPANY!MANUFACTURER
        End If
        Set Me.Datacompany.RowSource = RSTCOMPANY
        Datacompany.ListField = "MANUFACTURER"
        Datacompany.BoundColumn = "MANUFACTURER"
    End If
    Exit Sub
eRRhAND:
    MsgBox Err.Description

End Sub

Private Sub txtcompany_GotFocus()
    txtcompany.SelStart = 0
    txtcompany.SelLength = Len(txtcompany.Text)
End Sub

Private Sub txtcompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If chknewcomp.value = 0 And Datacompany.BoundText = "" Then
                If Datacompany.VisibleCount = 0 Then Exit Sub
'                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
'                Txtcompany.SetFocus
                Datacompany.SetFocus
                Exit Sub
            Else
                CmbPack.SetFocus
            End If
        Case vbKeyEscape
            txtcategory.SetFocus
    End Select

End Sub

Private Sub txtcompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_Click()

    txtcompany.Text = Datacompany.Text
    LBLDEALER2.Caption = txtcompany.Text

End Sub

Private Sub Datacompany_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtcompany.Text = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If
            If chknewcomp.value = 0 And Datacompany.BoundText = "" Then
                MsgBox "SELECT COMPANY NAME", vbOKOnly, "PRODUCT MASTER"
                txtcompany.SetFocus
                Exit Sub
            End If

            CmbPack.SetFocus
        Case vbKeyEscape
            txtcompany.SetFocus
    End Select
End Sub

Private Sub Datacompany_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Datacompany_GotFocus()
    FLAGCHANGE2.Caption = 1
    txtcompany.Text = LBLDEALER2.Caption
    Datacompany.Text = txtcompany.Text
    Call Datacompany_Click
End Sub

Private Sub Datacompany_LostFocus()
     FLAGCHANGE2.Caption = ""
End Sub

Private Sub txtcategory_LostFocus()
    If UCase(Datacategory.BoundText) = "SERVICE CHARGE" Then
        FrmeCompany.Visible = False
        
        CmbPack.Visible = False
        
        
    Else
        FrmeCompany.Visible = True
        CmbPack.Visible = True
        
        
        
    End If
End Sub
