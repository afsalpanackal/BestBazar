VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStockCorrect 
   Caption         =   "Stock Correction"
   ClientHeight    =   8370
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18300
   Icon            =   "frmStockCorrect.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8370
   ScaleWidth      =   18300
   Begin VB.CheckBox CHKCATEGORY2 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   11085
      TabIndex        =   18
      Top             =   90
      Width           =   1590
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
      Left            =   11055
      TabIndex        =   17
      Top             =   345
      Width           =   3225
   End
   Begin VB.CheckBox chkcategory 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   12840
      TabIndex        =   16
      Top             =   90
      Width           =   1410
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
      Left            =   15
      TabIndex        =   11
      Top             =   330
      Width           =   2625
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
      Left            =   2655
      TabIndex        =   10
      Top             =   330
      Width           =   1050
   End
   Begin VB.TextBox TxtItemcode 
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
      Left            =   3720
      TabIndex        =   9
      Top             =   330
      Width           =   1140
   End
   Begin VB.TextBox TxtTax 
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
      Left            =   4875
      TabIndex        =   8
      Top             =   330
      Width           =   855
   End
   Begin VB.TextBox TxtHSNCODE 
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
      Left            =   5745
      TabIndex        =   7
      Top             =   330
      Width           =   990
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
      Height          =   405
      Left            =   10455
      TabIndex        =   6
      Top             =   585
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1035
      Left            =   6750
      TabIndex        =   2
      Top             =   -45
      Width           =   2385
      Begin VB.OptionButton OptStock 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Stock Items Only"
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
         TabIndex        =   5
         Top             =   390
         Width           =   1935
      End
      Begin VB.OptionButton OptAll 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Display All"
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
         TabIndex        =   4
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton OptPC 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&Price Changing Items"
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
         TabIndex        =   3
         Top             =   660
         Width           =   2340
      End
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "Re- Load"
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
      Left            =   10455
      TabIndex        =   1
      Top             =   135
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdmsc 
      Height          =   7335
      Left            =   0
      TabIndex        =   0
      Top             =   1020
      Width           =   18285
      _ExtentX        =   32253
      _ExtentY        =   12938
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   19
      RowDividerStyle =   4
      AllowAddNew     =   -1  'True
      AllowDelete     =   -1  'True
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
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   11055
      TabIndex        =   19
      Top             =   690
      Width           =   3225
      _ExtentX        =   5689
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
   Begin VB.Label lblpvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   15930
      TabIndex        =   23
      Top             =   525
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   6
      Left            =   15900
      TabIndex        =   22
      Top             =   300
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Net Value"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   4
      Left            =   14310
      TabIndex        =   21
      Top             =   300
      Width           =   1185
   End
   Begin VB.Label lblnetvalue 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   390
      Left            =   14295
      TabIndex        =   20
      Top             =   525
      Width           =   1590
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
      Left            =   15
      TabIndex        =   15
      Top             =   75
      Width           =   3690
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
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   3720
      TabIndex        =   14
      Top             =   75
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Tax"
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
      Index           =   1
      Left            =   4875
      TabIndex        =   13
      Top             =   75
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "HSN Code"
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
      Index           =   5
      Left            =   5745
      TabIndex        =   12
      Top             =   75
      Width           =   990
   End
End
Attribute VB_Name = "frmStockCorrect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private adoGrid As ADODB.Recordset

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error GoTo ErrHand
    
    Dim db2 As New ADODB.Connection
    Dim DBPwd As String
    Dim DBPath As String
    Dim RSTITEMMAST As ADODB.Recordset
    Dim RstCircle As ADODB.Recordset
    Dim strConnect As String
    
    Set db2 = New ADODB.Connection
    DBPath = "D:\Tower\Analysis.db"
    '"\\192.168.1.3\data (d)\dbase1
    DBPwd = "donotopenme"
    
    Screen.MousePointer = vbHourglass
    If MsgBox("Are you sure...???", vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & DBPath & ";Jet OLEDB:Database Password=" & DBPwd
    db2.Open strConnect
    db2.CursorLocation = adUseClient
    
    db2.Execute "delete * from SP_DETAILS"
    
    Set RSTITEMMAST = New ADODB.Recordset
    Set RstCircle = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM SP_DETAILS ", db, adOpenStatic, adLockReadOnly, adCmdText
    RstCircle.Open "SELECT *  FROM SP_DETAILS ", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        RstCircle.AddNew
        RstCircle!MOB_CODE = RSTITEMMAST!MOB_CODE
        RstCircle!SP_NAME = RSTITEMMAST!SP_NAME
        RstCircle!SP_CIRCLE = Left(RSTITEMMAST!Circle, 25)
        RstCircle.Update
        RSTITEMMAST.MoveNext
    Loop
    RstCircle.Close
    Set RstCircle = Nothing
    
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    db2.Close
    Set db2 = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Updated succesfully", , "Cybersoft"
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number = -2147217865 Then
        MsgBox "Table Not Found. Please avail help from the vendor.", , "CyberSoft"
    Else
        MsgBox Err.Description, , "CyberSoft"
    End If
End Sub

Private Sub Form_Load()
    Set grdmsc.DataSource = Nothing
'    If adoGrid.State = 1 Then
'        adoGrid.Close
'        Set adoGrid = Nothing
'    End If
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        .Open "SELECT MOB_CODE, SP_NAME, CIRCLE_CODE, CIRCLE FROM SP_DETAILS order by mob_code", db, adOpenDynamic, adLockOptimistic
    End With
    Set grdmsc.DataSource = adoGrid
    grdmsc.Columns(0).Caption = "First 4 digit"
    grdmsc.Columns(1).Caption = "Service Provider"
    grdmsc.Columns(2).Caption = "Circle Code"
    grdmsc.Columns(3).Caption = "Circle Name"
    
    grdmsc.Columns(0).Width = 1500
    grdmsc.Columns(1).Width = 2500
    grdmsc.Columns(2).Width = 1200
    grdmsc.Columns(3).Width = 4000
    
'    On Error GoTo eRRHAND
'    Set GrdMSC.DataSource = Nothing
'    If PHYFLAG = True Then
'        PHY.Open "Select * FROM SP_DETAILS", db, adOpenStatic, adLockReadOnly
'        PHYFLAG = False
'    Else
'        PHY.Close
'        'PHY.Open "Select ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' AND ITEM_CODE <> '" & Trim(TXTITEMCODE.Text) & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
'        PHY.Open "Select ITEM_CODE, ITEM_NAME, CATEGORY FROM ITEMMAST WHERE ITEM_NAME Like '%" & Trim(Me.TXTITEM.Text) & "%' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly
'        PHYFLAG = False
'    End If
'    Set GrdMSC.DataSource = PHY
'    GrdMSC.Columns(0).Caption = "Code"
'    'GrdMSC.Columns(8).Caption = ""
'
'    GrdMSC.Columns(0).Width = 1000
'    GrdMSC.Columns(1).Width = 3800
'    GrdMSC.Columns(2).Width = 1200
    
    
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If adoGrid.State = 1 Then
        adoGrid.Close
        Set adoGrid = Nothing
    End If
End Sub

Private Sub grdmsc_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
                 
        Case 46
            If MsgBox("Are you sure you want to delete the entry...???", vbYesNo + vbDefaultButton2) = vbYes Then Exit Sub
            KeyCode = 0
    End Select
End Sub

Private Sub TxtCircle_Change()
    On Error GoTo ErrHand
    Set grdmsc.DataSource = Nothing
    If adoGrid.State = 1 Then
        adoGrid.Close
        Set adoGrid = Nothing
    End If
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        .Open "SELECT MOB_CODE, SP_NAME, CIRCLE_CODE, CIRCLE FROM SP_DETAILS WHERE CIRCLE Like '%" & TxtCircle.Text & "%' order by mob_code", db, adOpenDynamic, adLockOptimistic
    End With
    Set grdmsc.DataSource = adoGrid
    grdmsc.Columns(0).Caption = "First 4 digit"
    grdmsc.Columns(1).Caption = "Service Provider"
    grdmsc.Columns(2).Caption = "Circle Code"
    grdmsc.Columns(3).Caption = "Circle Name"
    
    grdmsc.Columns(0).Width = 1500
    grdmsc.Columns(1).Width = 2500
    grdmsc.Columns(2).Width = 1200
    grdmsc.Columns(3).Width = 4000
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TxtCircle_GotFocus()
    TxtCircle.SelStart = 0
    TxtCircle.SelLength = Len(TxtCircle.Text)
End Sub

Private Sub TxtCircle_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("]"), Asc("[")
            KeyAscii = 0
    End Select
End Sub

Private Sub Txtmobno_Change()
    On Error GoTo ErrHand
    Set grdmsc.DataSource = Nothing
    If adoGrid.State = 1 Then
        adoGrid.Close
        Set adoGrid = Nothing
    End If
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        .Open "SELECT MOB_CODE, SP_NAME, CIRCLE_CODE, CIRCLE FROM SP_DETAILS WHERE MOB_CODE Like '%" & Txtmobno.Text & "%' order by mob_code", db, adOpenDynamic, adLockOptimistic
    End With
    Set grdmsc.DataSource = adoGrid
    grdmsc.Columns(0).Caption = "First 4 digit"
    grdmsc.Columns(1).Caption = "Service Provider"
    grdmsc.Columns(2).Caption = "Circle Code"
    grdmsc.Columns(3).Caption = "Circle Name"
    
    grdmsc.Columns(0).Width = 1500
    grdmsc.Columns(1).Width = 2500
    grdmsc.Columns(2).Width = 1200
    grdmsc.Columns(3).Width = 4000
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub Txtmobno_GotFocus()
    Txtmobno.SelStart = 0
    Txtmobno.SelLength = Len(Txtmobno.Text)
End Sub

Private Sub Txtmobno_KeyPress(KeyAscii As Integer)
    
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

