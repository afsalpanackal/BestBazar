VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form Frmitemmerge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK MOVEMENT"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11340
   ClipControls    =   0   'False
   Icon            =   "frmitemmerge.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   11340
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
      Left            =   7395
      TabIndex        =   20
      Top             =   210
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
      Left            =   9630
      TabIndex        =   19
      Top             =   210
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
      Left            =   5775
      TabIndex        =   18
      Top             =   210
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
      Left            =   5775
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   1620
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
      Left            =   4095
      TabIndex        =   11
      Top             =   7980
      Width           =   1440
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
      Height          =   300
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1620
      Width           =   4050
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
      Left            =   9870
      TabIndex        =   4
      Top             =   7980
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
   Begin MSDataGridLib.DataGrid grd1IN 
      Height          =   2730
      Left            =   30
      TabIndex        =   12
      Top             =   2205
      Width           =   5505
      _ExtentX        =   9710
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
   Begin MSDataGridLib.DataGrid grd1OUT 
      Height          =   2730
      Left            =   30
      TabIndex        =   13
      Top             =   5205
      Width           =   5505
      _ExtentX        =   9710
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
   Begin MSDataGridLib.DataGrid grd2IN 
      Height          =   2730
      Left            =   5775
      TabIndex        =   15
      Top             =   2205
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
      Left            =   5775
      TabIndex        =   16
      Top             =   5205
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
      Left            =   5775
      TabIndex        =   21
      Top             =   570
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
   Begin VB.Frame frmunbill 
      BackColor       =   &H00FFC0C0&
      Height          =   690
      Left            =   5760
      TabIndex        =   27
      Top             =   7905
      Visible         =   0   'False
      Width           =   2565
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
         TabIndex        =   29
         Top             =   180
         Width           =   1875
      End
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
      TabIndex        =   26
      Top             =   8040
      Width           =   3945
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
      Left            =   5775
      TabIndex        =   25
      Top             =   4950
      Width           =   5475
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
      Index           =   2
      Left            =   30
      TabIndex        =   24
      Top             =   4950
      Width           =   5505
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      DrawMode        =   9  'Not Mask Pen
      X1              =   5625
      X2              =   5625
      Y1              =   210
      Y2              =   7935
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
      Left            =   5790
      TabIndex        =   23
      Top             =   15
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
      Left            =   9645
      TabIndex        =   22
      Top             =   15
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
      Left            =   5775
      TabIndex        =   14
      Top             =   1950
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
      TabIndex        =   6
      Top             =   8910
      Visible         =   0   'False
      Width           =   1740
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
      Index           =   0
      Left            =   30
      TabIndex        =   5
      Top             =   1950
      Width           =   5505
   End
End
Attribute VB_Name = "Frmitemmerge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RSTREP As New ADODB.Recordset
Dim RSTREP2 As New ADODB.Recordset
Private adoGridIN1 As New ADODB.Recordset
Private adoGridIN2 As New ADODB.Recordset
Private adoGridOUT1 As New ADODB.Recordset
Private adoGridOUT2 As New ADODB.Recordset

Private Sub cmdexit_Click()
    Unload Me
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
    If MsgBox("Are you sure you want to merge the item " & DataList2.text & " with " & DataList1.text & ". The " & DataList2.text & " will be deleted after merging.", vbYesNo + vbDefaultButton2, "ITEM MERGE....") = vbNo Then Exit Sub
    
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo ErrHand
    db.Execute "Update RTRXFILE set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & DataList2.BoundText & "' "
    db.Execute "Update TRXFILE set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & DataList2.BoundText & "' "
    db.Execute "Update TRXFORMULASUB set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & DataList2.BoundText & "' "
    db.Execute "Update TRXFORMULASUB set FOR_NAME = '" & DataList1.BoundText & "' where FOR_NAME = '" & DataList2.BoundText & "' "
    db.Execute "Update TRXFORMULAMAST set ITEM_CODE = '" & DataList1.BoundText & "' where ITEM_CODE = '" & DataList2.BoundText & "' "
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
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
    rststock.Open "SELECT * from RTRXFILE where RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing

    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFILE where ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing

    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing

    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULASUB where FOR_NAME = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing

    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from TRXFORMULAMAST where ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    If Not (rststock.EOF And rststock.BOF) Then
        MsgBox "Cannot Delete " & DataList2.text & " Since Transactions is Available", vbCritical, "Item Merge"
        rststock.Close
        Set rststock = Nothing
        Exit Sub
    End If
    rststock.Close
    Set rststock = Nothing
    
    
    
    'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "'")
    db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & DataList2.BoundText & "'")
    db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & DataList2.BoundText & "'")
    
    adoGridIN1.Close
    adoGridOUT1.Close
    Set adoGridIN1 = Nothing
    Set adoGridOUT1 = Nothing
    Call tXTMEDICINE_Change
    
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
            'GRDSTOCK.SetFocus
            'DataList2.SetFocus
        Case vbKeyEscape
            TxtItemName.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    
    'Me.Height = 9990
    'Me.Width = 18555
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If RSTREP.State = 1 Then RSTREP.Close
    If RSTREP2.State = 1 Then RSTREP2.Close
    If adoGridIN1.State = 1 Then
        adoGridIN1.Close
        Set adoGridIN1 = Nothing
    End If
    If adoGridIN2.State = 1 Then
        adoGridIN2.Close
        Set adoGridIN2 = Nothing
    End If
    If adoGridOUT1.State = 1 Then
        adoGridOUT1.Close
        Set adoGridOUT1 = Nothing
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
        
    LBLHEAD(0).Caption = "INWARD DETAILS OF " & DataList2.text & " (" & DataList2.BoundText & ")"
    LBLHEAD(2).Caption = "OUTWARD DETAILS OF " & DataList2.text & " (" & DataList2.BoundText & ")"
    Set grd1IN.DataSource = Nothing
    Set adoGridIN1 = New ADODB.Recordset
    With adoGridIN1
        .CursorLocation = adUseClient
        .Open "SELECT QTY, VCH_DATE, PINV, VCH_NO, TRX_TYPE FROM RTRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    End With
    Set grd1IN.DataSource = adoGridIN1
    
    grd1IN.Columns(0).Caption = "QTY"
    grd1IN.Columns(1).Caption = "DATE"
    grd1IN.Columns(2).Caption = "INV NO"
    grd1IN.Columns(3).Caption = "REF NO"
    grd1IN.Columns(4).Caption = "TYPE"
    
    grd1IN.Columns(0).Width = 700
    grd1IN.Columns(1).Width = 1200
    grd1IN.Columns(2).Width = 1200
    grd1IN.Columns(3).Width = 1200
    grd1IN.Columns(4).Width = 1000
    
    Set grd1OUT.DataSource = Nothing
    Set adoGridOUT1 = New ADODB.Recordset
    With adoGridOUT1
        .CursorLocation = adUseClient
        .Open "SELECT QTY, VCH_DATE, VCH_NO, TRX_TYPE FROM TRXFILE WHERE  ITEM_CODE = '" & DataList2.BoundText & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ORDER BY VCH_DATE DESC, VCH_NO", db, adOpenStatic, adLockReadOnly
    End With
    Set grd1OUT.DataSource = adoGridOUT1
    
    grd1OUT.Columns(0).Caption = "QTY"
    grd1OUT.Columns(1).Caption = "DATE"
    grd1OUT.Columns(2).Caption = "BILL NO"
    grd1OUT.Columns(3).Caption = "TYPE"
    
    grd1OUT.Columns(0).Width = 700
    grd1OUT.Columns(1).Width = 1200
    grd1OUT.Columns(2).Width = 1200
    grd1OUT.Columns(3).Width = 1000
    
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

