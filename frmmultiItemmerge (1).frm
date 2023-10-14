VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmitemmergeMulti 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK MOVEMENT (Multiple Items)"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14595
   ClipControls    =   0   'False
   Icon            =   "frmmultiItemmerge.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14595
   Begin VB.Frame frmunbill 
      BackColor       =   &H00FFC0C0&
      Height          =   690
      Left            =   6285
      TabIndex        =   26
      Top             =   1350
      Visible         =   0   'False
      Width           =   2565
      Begin VB.CheckBox chkonlyunbill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Only Un Bill Items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   28
         Top             =   420
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CheckBox Chkunbill 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Show Un Bill items"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   45
         TabIndex        =   27
         Top             =   180
         Width           =   1875
      End
   End
   Begin VB.CommandButton Cmdremove 
      Caption         =   "&Remove"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4515
      TabIndex        =   25
      Top             =   1665
      Width           =   1005
   End
   Begin VB.TextBox LBLITEMCODE 
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
      Height          =   360
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   24
      Top             =   1665
      Width           =   3390
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3465
      TabIndex        =   22
      Top             =   1665
      Width           =   1020
   End
   Begin VB.TextBox tXTMEDICINE2 
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
      Left            =   10605
      TabIndex        =   16
      Top             =   255
      Width           =   2220
   End
   Begin VB.TextBox tXTCODE2 
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
      Left            =   12840
      TabIndex        =   15
      Top             =   255
      Width           =   1590
   End
   Begin VB.TextBox TxtItemName2 
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
      Left            =   8985
      TabIndex        =   14
      Top             =   255
      Width           =   1605
   End
   Begin VB.TextBox LBLITEMCODE2 
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
      Left            =   8985
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1665
      Width           =   4050
   End
   Begin VB.CommandButton cmdstkcrct 
      Caption         =   "Item Merge >>"
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
      Left            =   7410
      TabIndex        =   9
      Top             =   7980
      Width           =   1440
   End
   Begin VB.TextBox TxtItemName 
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
      Left            =   30
      TabIndex        =   0
      Top             =   210
      Width           =   1605
   End
   Begin VB.TextBox tXTCODE 
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
      Left            =   3885
      TabIndex        =   2
      Top             =   210
      Width           =   1635
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
      Left            =   1650
      TabIndex        =   1
      Top             =   210
      Width           =   2220
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
      Left            =   13080
      TabIndex        =   4
      Top             =   8025
      Width           =   1380
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1035
      Left            =   30
      TabIndex        =   3
      Top             =   570
      Width           =   5490
      _ExtentX        =   9684
      _ExtentY        =   1826
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
   Begin MSDataGridLib.DataGrid grd2IN 
      Height          =   2730
      Left            =   8985
      TabIndex        =   11
      Top             =   2250
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   4815
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
   Begin MSDataGridLib.DataGrid grd2OUT 
      Height          =   2730
      Left            =   8985
      TabIndex        =   12
      Top             =   5250
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   4815
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
      Height          =   1035
      Left            =   8985
      TabIndex        =   17
      Top             =   615
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   1826
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
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   5865
      Left            =   15
      TabIndex        =   23
      Top             =   2070
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   10345
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   400
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "Note: Left side item will be merged to the right side. Further, the leftside item will be deleted."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   465
      Left            =   60
      TabIndex        =   21
      Top             =   8040
      Width           =   6990
      WordWrap        =   -1  'True
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
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
      Index           =   3
      Left            =   8985
      TabIndex        =   20
      Top             =   4995
      Width           =   5475
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      DrawMode        =   9  'Not Mask Pen
      X1              =   8895
      X2              =   8895
      Y1              =   225
      Y2              =   7950
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Height          =   225
      Index           =   1
      Left            =   9000
      TabIndex        =   19
      Top             =   60
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Height          =   210
      Index           =   0
      Left            =   12855
      TabIndex        =   18
      Top             =   60
      Width           =   1170
   End
   Begin VB.Label LBLHEAD 
      BackColor       =   &H00000000&
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
      Index           =   1
      Left            =   8985
      TabIndex        =   10
      Top             =   1995
      Width           =   5460
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Code"
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
      Height          =   210
      Index           =   13
      Left            =   3900
      TabIndex        =   8
      Top             =   15
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Height          =   225
      Index           =   12
      Left            =   45
      TabIndex        =   7
      Top             =   15
      Width           =   1185
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "LOOSE QTY"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   345
      Index           =   2
      Left            =   3390
      TabIndex        =   6
      Top             =   8970
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Label LblLoose 
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
      ForeColor       =   &H000040C0&
      Height          =   360
      Left            =   5265
      TabIndex        =   5
      Top             =   8910
      Visible         =   0   'False
      Width           =   1740
   End
End
Attribute VB_Name = "FrmitemmergeMulti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSTREP As New ADODB.Recordset
Dim RSTREP2 As New ADODB.Recordset
Private adoGridIN2 As New ADODB.Recordset
Private adoGridOUT2 As New ADODB.Recordset

Private Sub CMDADD_Click()
    Dim RSTITEM As ADODB.Recordset
                
    Dim i As Integer
    For i = 1 To GRDSTOCK.rows - 1
        If Trim(GRDSTOCK.TextMatrix(i, 1)) = DataList2.BoundText Then
            MsgBox "This Item Already added.", , "Item Merge"
            TxtItemName.SetFocus
            Exit Sub
        End If
    Next i
    
    On Error GoTo ErrHand
    Set RSTITEM = New ADODB.Recordset
    RSTITEM.Open "SELECT * FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "' ", db, adOpenStatic, adLockReadOnly
    If Not (RSTITEM.EOF And RSTITEM.BOF) Then
        GRDSTOCK.rows = GRDSTOCK.rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 0) = GRDSTOCK.rows - 1
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 1) = RSTITEM!ITEM_CODE
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 2) = RSTITEM!ITEM_NAME
        GRDSTOCK.TextMatrix(GRDSTOCK.rows - 1, 3) = RSTITEM!CLOSE_QTY
    End If
    RSTITEM.Close
    Set RSTITEM = Nothing
            
    Exit Sub
ErrHand:
    MsgBox err.Description
End Sub

Private Sub cmdadd_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Cmdremove_Click()
    
    Dim selrow As Integer
    Dim i As Long
    selrow = GRDSTOCK.Row
    For i = selrow To GRDSTOCK.rows - 2
        GRDSTOCK.TextMatrix(selrow, 0) = i
        GRDSTOCK.TextMatrix(selrow, 1) = GRDSTOCK.TextMatrix(i + 1, 1)
        GRDSTOCK.TextMatrix(selrow, 2) = GRDSTOCK.TextMatrix(i + 1, 2)
        GRDSTOCK.TextMatrix(selrow, 3) = GRDSTOCK.TextMatrix(i + 1, 3)
        GRDSTOCK.TextMatrix(selrow, 4) = GRDSTOCK.TextMatrix(i + 1, 4)
        
        selrow = selrow + 1
    Next i
    GRDSTOCK.rows = GRDSTOCK.rows - 1
    GRDSTOCK.SetFocus
    Exit Sub
   
ErrHand:
    MsgBox err.Description
End Sub

Private Sub Cmdremove_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub CmdStkCrct_Click()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If MDIMAIN.StatusBar.Panels(9).text = "Y" Then Exit Sub
    
    If DataList2.BoundText = "" Then Exit Sub
    If DataList1.BoundText = "" Then Exit Sub
    If DataList2.BoundText = DataList1.BoundText Then
        MsgBox "Items are same", , "Item Merge"
        Exit Sub
    End If
    If MsgBox("Are you sure you want to merge the selected items with " & DataList1.text & ". The selected items will be deleted after merging.", vbYesNo + vbDefaultButton2, "ITEM MERGE....") = vbNo Then Exit Sub
    
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ErrHand
    
    Dim i As Integer
    For i = 1 To GRDSTOCK.rows - 1
        db.Execute "Update RTRXFILE set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' "
        db.Execute "Update TRXFILE set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' "
        db.Execute "Update TRXFORMULASUB set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' "
        db.Execute "Update TRXFORMULASUB set FOR_NAME = '" & DataList1.BoundText & "' where FOR_NAME = '" & GRDSTOCK.TextMatrix(i, 1) & "' "
        db.Execute "Update TRXFORMULAMAST set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "' "
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            Set RSTITEMMAST = New ADODB.Recordset
            RSTITEMMAST.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & DataList1.BoundText & "'", db, adOpenStatic, adLockPessimistic, adCmdText
            If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                RSTITEMMAST!OPEN_QTY = IIf(IsNull(RSTITEMMAST!OPEN_QTY), 0, RSTITEMMAST!OPEN_QTY) + IIf(IsNull(rststock!OPEN_QTY), 0, rststock!OPEN_QTY)
                RSTITEMMAST!OPEN_VAL = IIf(IsNull(RSTITEMMAST!OPEN_VAL), 0, RSTITEMMAST!OPEN_VAL) + IIf(IsNull(rststock!OPEN_VAL), 0, rststock!OPEN_VAL)
                
                RSTITEMMAST!CLOSE_QTY = IIf(IsNull(RSTITEMMAST!CLOSE_QTY), 0, RSTITEMMAST!CLOSE_QTY) + IIf(IsNull(rststock!CLOSE_QTY), 0, rststock!CLOSE_QTY)
                RSTITEMMAST!CLOSE_VAL = IIf(IsNull(RSTITEMMAST!CLOSE_VAL), 0, RSTITEMMAST!CLOSE_VAL) + IIf(IsNull(rststock!CLOSE_VAL), 0, rststock!CLOSE_VAL)
                RSTITEMMAST.Update
            End If
            RSTITEMMAST.Close
            Set RSTITEMMAST = Nothing
        End If
        rststock.Close
        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
            rststock.Close
            Set rststock = Nothing
            Exit Sub
        End If
        rststock.Close
        Set rststock = Nothing
    
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
            rststock.Close
            Set rststock = Nothing
            Exit Sub
        End If
        rststock.Close
        Set rststock = Nothing
    
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from TRXFORMULASUB where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
            rststock.Close
            Set rststock = Nothing
            Exit Sub
        End If
        rststock.Close
        Set rststock = Nothing
    
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
            rststock.Close
            Set rststock = Nothing
            Exit Sub
        End If
        rststock.Close
        Set rststock = Nothing
    
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
            rststock.Close
            Set rststock = Nothing
            Exit Sub
        End If
        rststock.Close
        Set rststock = Nothing
        
        
        
        'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'")
        db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'")
        db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & GRDSTOCK.TextMatrix(i, 1) & "'")
    Next i
            
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.rows = 1
    TxtItemName.SetFocus
    
'    Call DataList2_Click
'    Call DataList1_Click
    
    Exit Sub
   
ErrHand:
    MsgBox err.Description
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            cmdadd.SetFocus
            'GRDSTOCK.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    
    'Me.Height = 9990
    'Me.Width = 18555
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "QTY"
    
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 1500
    GRDSTOCK.ColWidth(2) = 5000
    GRDSTOCK.ColWidth(3) = 1200
    
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    
    
    Me.Left = 0
    Me.Top = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RSTREP.State = 1 Then RSTREP.Close
    If RSTREP2.State = 1 Then RSTREP2.Close
    If adoGridIN2.State = 1 Then
        adoGridIN2.Close
        Set adoGridIN2 = Nothing
    End If
    If adoGridOUT2.State = 1 Then
        adoGridOUT2.Close
        Set adoGridOUT2 = Nothing
    End If
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub


Private Sub Label2_DblClick()
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    If frmunbill.Visible = False Then
        frmunbill.Visible = True
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
    Else
        frmunbill.Visible = False
        chkunbill.Value = 0
        chkonlyunbill.Value = 0
    End If
End Sub

Private Sub Chkunbill_Click()
    If chkunbill.Value = 1 Then
        chkonlyunbill.Visible = True
        chkonlyunbill.Value = 0
    Else
        chkonlyunbill.Visible = False
        chkonlyunbill.Value = 0
    End If
End Sub

Private Sub TXTCODE_Change()
    On Error GoTo ErrHand
    If RSTREP.State = 1 Then RSTREP.Close
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE UN_BILL = 'Y' AND ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_CODE Like '" & TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If
    
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo ErrHand
    If RSTREP.State = 1 Then RSTREP.Close
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE UN_BILL = 'Y' AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If
    
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub
Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(tXTMEDICINE.text) = "" Then
                TxtCode.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList2_Click()
    LBLITEMCODE.text = DataList2.BoundText
    Call Fillgrid
    'Call Fillgrid2
    Screen.MousePointer = vbNormal
    ''''''''LBLBALANCE.Caption = Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption)
End Sub

Private Function Fillgrid()
    
    If DataList2.BoundText = "" Then Exit Function
    On Error GoTo ErrHand

    
    Screen.MousePointer = vbHourglass
        
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Function Fillgrid2()
    
    If DataList1.BoundText = "" Then Exit Function
    On Error GoTo ErrHand

    
    Screen.MousePointer = vbHourglass
    LBLHEAD(1).Caption = "INWARD DETAILS OF " & DataList1.text & " (" & DataList1.BoundText & ")"
    LBLHEAD(3).Caption = "OUTWARD DETAILS OF " & DataList1.text & " (" & DataList1.BoundText & ")"
    
    Set grd2IN.DataSource = Nothing
    Set adoGridIN2 = New ADODB.Recordset
    With adoGridIN2
        .CursorLocation = adUseClient
        .Open "SELECT QTY, VCH_DATE, PINV, VCH_NO, TRX_TYPE FROM RTRXFILE WHERE  ITEM_CODE = '" & DataList1.BoundText & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    End With
    Set grd2IN.DataSource = adoGridIN2
    
    grd2IN.Columns(0).Caption = "QTY"
    grd2IN.Columns(1).Caption = "DATE"
    grd2IN.Columns(2).Caption = "INV NO"
    grd2IN.Columns(3).Caption = "REF NO"
    grd2IN.Columns(4).Caption = "TYPE"
    
    grd2IN.Columns(0).Width = 700
    grd2IN.Columns(1).Width = 1200
    grd2IN.Columns(2).Width = 1200
    grd2IN.Columns(3).Width = 1200
    grd2IN.Columns(4).Width = 1000
    
    Set grd2OUT.DataSource = Nothing
    Set adoGridOUT2 = New ADODB.Recordset
    With adoGridOUT2
        .CursorLocation = adUseClient
        .Open "SELECT QTY, VCH_DATE, VCH_NO, TRX_TYPE FROM TRXFILE WHERE  ITEM_CODE = '" & DataList1.BoundText & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    End With
    Set grd2OUT.DataSource = adoGridOUT2
    
    grd2OUT.Columns(0).Caption = "QTY"
    grd2OUT.Columns(1).Caption = "DATE"
    grd2OUT.Columns(2).Caption = "BILL NO"
    grd2OUT.Columns(3).Caption = "TYPE"
    
    grd2OUT.Columns(0).Width = 700
    grd2OUT.Columns(1).Width = 1200
    grd2OUT.Columns(2).Width = 1200
    grd2OUT.Columns(3).Width = 1000
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.text)
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(TxtCode.text) = "" Then
                TxtItemName.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select

End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtItemName_Change()
    
    On Error GoTo ErrHand
    If RSTREP.State = 1 Then RSTREP.Close
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE  UN_BILL = 'Y' AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.text & "%' AND ITEM_CODE Like '%" & Me.TxtCode.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If
    
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtItemName_GotFocus()
    TxtItemName.SelStart = 0
    TxtItemName.SelLength = Len(TxtItemName.text)
End Sub

Private Sub TxtItemName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            If Trim(TxtItemName.text) = "" Then
                tXTMEDICINE.SetFocus
            Else
                DataList2.SetFocus
            End If
        Case vbKeyEscape
            Call cmdexit_Click
    End Select

End Sub

Private Sub TxtItemName_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtItemName2_Change()
    
    On Error GoTo ErrHand
    If RSTREP2.State = 1 Then RSTREP2.Close
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE UN_BILL = 'Y' AND ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP2.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If
    
    Set Me.DataList1.RowSource = RSTREP2
    DataList1.ListField = "ITEM_NAME"
    DataList1.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP2.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtItemName2_GotFocus()
    TxtItemName2.SelStart = 0
    TxtItemName2.SelLength = Len(TxtItemName2.text)
End Sub

Private Sub TxtItemName2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList1.VisibleCount = 0 Then Exit Sub
            If Trim(TxtItemName2.text) = "" Then
                tXTMEDICINE2.SetFocus
            Else
                DataList1.SetFocus
            End If
    End Select

End Sub

Private Sub TxtItemName2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub tXTMEDICINE2_Change()
    
    On Error GoTo ErrHand
    If RSTREP2.State = 1 Then RSTREP2.Close
    
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE UN_BILL = 'Y' AND ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP2.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_NAME Like '" & Me.TxtItemName2.text & "%' AND ITEM_NAME Like '%" & Me.tXTMEDICINE2.text & "%' AND ITEM_CODE Like '%" & Me.tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If

    Set Me.DataList1.RowSource = RSTREP2
    DataList1.ListField = "ITEM_NAME"
    DataList1.BoundColumn = "ITEM_CODE"
    
    
    Exit Sub
'RSTREP2.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub


Private Sub tXTMEDICINE2_GotFocus()
    tXTMEDICINE2.SelStart = 0
    tXTMEDICINE2.SelLength = Len(tXTMEDICINE2.text)
End Sub

Private Sub tXTMEDICINE2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList1.VisibleCount = 0 Then Exit Sub
            If Trim(tXTMEDICINE2.text) = "" Then
                tXTCODE2.SetFocus
            Else
                DataList1.SetFocus
            End If
        Case vbKeyEscape
            TxtItemName2.SetFocus
    End Select

End Sub

Private Sub tXTMEDICINE2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TXTCODE2_Change()
    On Error GoTo ErrHand
    If RSTREP2.State = 1 Then RSTREP2.Close
    
    If chkunbill.Value = 1 And chkonlyunbill.Value = 1 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE UN_BILL = 'Y' AND ITEM_CODE Like '" & tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    ElseIf chkunbill.Value = 1 And chkonlyunbill.Value = 0 Then
        RSTREP2.Open "Select * From ITEMMAST  WHERE ITEM_CODE Like '" & tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    Else
        RSTREP2.Open "Select * From ITEMMAST  WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ITEM_CODE Like '" & tXTCODE2.text & "%' AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
    End If
    
    Set Me.DataList1.RowSource = RSTREP2
    DataList1.ListField = "ITEM_NAME"
    DataList1.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP2.Close
'TMPFLAG = False
ErrHand:
    MsgBox err.Description
End Sub

Private Sub TxtCode2_GotFocus()
    tXTCODE2.SelStart = 0
    tXTCODE2.SelLength = Len(tXTCODE2.text)
End Sub

Private Sub TxtCode2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList1.VisibleCount = 0 Then Exit Sub
            If Trim(tXTCODE2.text) = "" Then
                TxtItemName2.SetFocus
            Else
                DataList1.SetFocus
            End If
        Case vbKeyEscape
            tXTMEDICINE2.SetFocus
    End Select

End Sub

Private Sub TxtCode2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub DataList1_Click()
    LBLITEMCODE2.text = DataList1.BoundText
    Call Fillgrid2
    'Call Fillgrid2
    Screen.MousePointer = vbNormal
    ''''''''LBLBALANCE.Caption = Val(LBLINWARD.Caption) - Val(LBLOUTWARD.Caption)
End Sub

