VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmLocSearch 
   BackColor       =   &H00E8DFEC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Price Analysis"
   ClientHeight    =   8280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18705
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Frmlocsearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8280
   ScaleWidth      =   18705
   Begin VB.CommandButton CmdTranslate 
      Caption         =   "Apply Translation to Selected Items"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   14280
      TabIndex        =   37
      Top             =   1755
      Width           =   1920
   End
   Begin VB.TextBox txtspecs 
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
      Left            =   6780
      TabIndex        =   4
      Top             =   270
      Width           =   1740
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   16830
      TabIndex        =   30
      Top             =   465
      Visible         =   0   'False
      Width           =   1755
      Begin VB.CheckBox chkhidecomp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide Company"
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
         Left            =   40
         TabIndex        =   35
         Top             =   1125
         Width           =   1635
      End
      Begin VB.CheckBox chkhidecat 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide Category"
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
         Left            =   40
         TabIndex        =   34
         Top             =   885
         Width           =   1635
      End
      Begin VB.CheckBox chklwp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide LW Price"
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
         Left            =   40
         TabIndex        =   33
         Top             =   645
         Width           =   1635
      End
      Begin VB.CheckBox chkvp 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide VP"
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
         Left            =   40
         TabIndex        =   32
         Top             =   405
         Width           =   1635
      End
      Begin VB.CheckBox chkws 
         Appearance      =   0  'Flat
         BackColor       =   &H00D5F2E6&
         Caption         =   "Hide WS"
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
         Left            =   40
         TabIndex        =   31
         Top             =   165
         Width           =   1635
      End
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
      Left            =   4920
      TabIndex        =   3
      Top             =   270
      Width           =   1845
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
      Left            =   3750
      TabIndex        =   2
      Top             =   270
      Width           =   1140
   End
   Begin VB.CheckBox chkcategory 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Category"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   12735
      TabIndex        =   14
      Top             =   45
      Width           =   1410
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
      Left            =   10950
      TabIndex        =   15
      Top             =   300
      Width           =   3225
   End
   Begin VB.CheckBox CHKCATEGORY2 
      BackColor       =   &H00E8DFEC&
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   10980
      TabIndex        =   13
      Top             =   45
      Width           =   1590
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
      Left            =   8535
      TabIndex        =   10
      Top             =   975
      Width           =   1215
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
      Left            =   2685
      TabIndex        =   1
      Top             =   270
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFC0C0&
      Height          =   1035
      Left            =   8535
      TabIndex        =   6
      Top             =   -75
      Width           =   2385
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
         TabIndex        =   9
         Top             =   660
         Width           =   2340
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
         TabIndex        =   7
         Top             =   120
         Value           =   -1  'True
         Width           =   1335
      End
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
         TabIndex        =   8
         Top             =   390
         Width           =   1935
      End
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
      Left            =   45
      TabIndex        =   0
      Top             =   270
      Width           =   2625
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
      Left            =   9780
      TabIndex        =   12
      Top             =   975
      Width           =   1125
   End
   Begin MSDataListLib.DataList DataList2 
      Height          =   1620
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   8475
      _ExtentX        =   14949
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
   Begin VB.Frame Frame1 
      Height          =   5985
      Left            =   45
      TabIndex        =   11
      Top             =   2295
      Width           =   18645
      Begin MSDataGridLib.DataGrid grdmsc 
         Height          =   5820
         Left            =   15
         TabIndex        =   18
         Top             =   120
         Width           =   18585
         _ExtentX        =   32782
         _ExtentY        =   10266
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
   End
   Begin MSComCtl2.DTPicker DTFROM 
      Height          =   375
      Left            =   8580
      TabIndex        =   17
      Top             =   1575
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarForeColor=   0
      CalendarTitleForeColor=   0
      CalendarTrailingForeColor=   255
      CheckBox        =   -1  'True
      Format          =   112001025
      CurrentDate     =   40498
   End
   Begin MSDataListLib.DataList DataList1 
      Height          =   780
      Left            =   10950
      TabIndex        =   16
      Top             =   645
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Specifications"
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
      Left            =   6780
      TabIndex        =   36
      Top             =   15
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Location"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   15
      Width           =   1845
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
      Left            =   3750
      TabIndex        =   28
      Top             =   15
      Width           =   1140
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
      Left            =   45
      TabIndex        =   27
      Top             =   15
      Width           =   3690
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
      Left            =   10965
      TabIndex        =   26
      Top             =   1905
      Width           =   1590
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
      Left            =   10980
      TabIndex        =   25
      Top             =   1680
      Width           =   1185
   End
   Begin VB.Label LBLDEALER2 
      Height          =   315
      Left            =   0
      TabIndex        =   24
      Top             =   810
      Visible         =   0   'False
      Width           =   1620
   End
   Begin VB.Label FLAGCHANGE2 
      Height          =   315
      Left            =   0
      TabIndex        =   23
      Top             =   450
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "OP. Stock Entry Date"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   315
      Index           =   3
      Left            =   8520
      TabIndex        =   22
      Top             =   1380
      Width           =   1890
      WordWrap        =   -1  'True
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
      Left            =   12570
      TabIndex        =   21
      Top             =   1680
      Width           =   1500
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
      Left            =   12600
      TabIndex        =   20
      Top             =   1905
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Part"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   2
      Left            =   60
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   1710
   End
End
Attribute VB_Name = "FrmLocSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim REPFLAG As Boolean 'REP
Dim MFG_REC As New ADODB.Recordset
Dim CAT_REC As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim PHY_FLAG As Boolean 'REP
Dim PHY_REC As New ADODB.Recordset
Dim frmloadflag As Boolean
Private adoGrid As ADODB.Recordset

Private Sub CHKCATEGORY_Click()
    CHKCATEGORY2.Value = 0
End Sub

Private Sub CHKCATEGORY2_Click()
    chkcategory.Value = 0
End Sub

Private Sub chkhidecat_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo eRRhAND
    If chkhidecat.Value = 1 Then
        db.Execute "Update COMPINFO set hide_category = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 18) = ""
'        GRDSTOCK.ColWidth(18) = 0
    Else
        db.Execute "Update COMPINFO set hide_category = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 18) = "Category"
'        GRDSTOCK.ColWidth(18) = 1000
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub chkhidecomp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo eRRhAND
    If chkhidecomp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_company = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 19) = ""
'        GRDSTOCK.ColWidth(19) = 0
    Else
        db.Execute "Update COMPINFO set hide_company = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 19) = "Company"
'        GRDSTOCK.ColWidth(19) = 1000
    End If
    
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub chklwp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo eRRhAND
    If chklwp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_lwp = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 17) = ""
'        GRDSTOCK.ColWidth(17) = 0
    Else
        db.Execute "Update COMPINFO set hide_lwp = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
'        GRDSTOCK.ColWidth(17) = 1000
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub chkvp_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo eRRhAND
    If chkvp.Value = 1 Then
        db.Execute "Update COMPINFO set hide_van = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 9) = ""
'        GRDSTOCK.ColWidth(9) = 0
    Else
        db.Execute "Update COMPINFO set hide_van = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 9) = "VP"
'        GRDSTOCK.ColWidth(9) = 1000
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub chkws_Click()
    If frmloadflag = True Then Exit Sub
    On Error GoTo eRRhAND
    If chkws.Value = 1 Then
        db.Execute "Update COMPINFO set hide_ws = 'Y' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 8) = ""
'        GRDSTOCK.ColWidth(8) = 0
    Else
        db.Execute "Update COMPINFO set hide_ws = 'N' where COMP_CODE = '001' "
'        GRDSTOCK.TextMatrix(0, 8) = "WS"
'        GRDSTOCK.ColWidth(8) = 1000
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdLoad_Click()
    Call Fillgrid
End Sub

Private Sub CmdTranslate_Click()
    Dim strURL As String
    Dim strResponse As String
    Dim XMLHttpRequest As XMLHTTP60
    Dim result() As String
    Dim eng_word As String
    Dim arab_word As String
    Dim i As Long
    'strURL = "http://nimbusit.co.in/api/swsendSingle.asp?username=t1gyanendramix&password=Gyani@123&sender=GYANAS&sendto=919072999927&message=Test SMS, HOW ARE YOU&entityID=1701160224350444363"
                            
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    
    Screen.MousePointer = vbHourglass
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
            If OptStock.Value = True Then
                .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            ElseIf OptPC.Value = True Then
                .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(ITEM_MAL) OR ITEM_MAL = '') and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If CHKCATEGORY2.Value = 1 Then
                If OptStock.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(ITEM_MAL) OR ITEM_MAL = '') and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
            Else
                If OptStock.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
                ElseIf OptPC.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(ITEM_MAL) OR ITEM_MAL = '') and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, P_RETAIL, P_WS, P_VAN, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenForwardOnly
                End If
    
            End If
        End If
        If .RecordCount > 0 Then
            MDIMAIN.vbalProgressBar1.Max = .RecordCount
        Else
            MDIMAIN.vbalProgressBar1.Max = 100
        End If
        i = 1
        Do Until .EOF
            
            If IsConnected = False Then
                Screen.MousePointer = vbNormal
                MsgBox "You need an internet Connection for translation.", vbOKOnly, "EzBiz"
                .Close
                Set adoGrid = Nothing
                Exit Sub
            End If
            
            eng_word = !ITEM_NAME
            arab_word = ""
            strURL = "https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=ml&dt=t&dt=bd&dj=1&q=" & eng_word
            
            Set XMLHttpRequest = New MSXML2.XMLHTTP60
            XMLHttpRequest.Open "GET", strURL, False
            XMLHttpRequest.setRequestHeader "Content-Type", "text/xml"
            XMLHttpRequest.send
        
            
            result() = Split(XMLHttpRequest.responseText, ",")
            arab_word = result(0)
            If arab_word <> "" Then
                arab_word = Mid(arab_word, 25)
                arab_word = Mid(arab_word, 1, Len(arab_word) - 1)
            End If
            If arab_word <> "" Then
                db.Execute "Update itemmast set ITEM_MAL ='" & arab_word & "' where ITEM_CODE = '" & !ITEM_CODE & "' "
            End If
            Set XMLHttpRequest = Nothing
            
            MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
            MDIMAIN.vbalProgressBar1.Text = i & "out of " & .RecordCount
            i = i + 1
            .MoveNext
        Loop
        .Close
        Set adoGrid = Nothing
    End With
    
    Call Fillgrid
    MDIMAIN.vbalProgressBar1.ShowText = False
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, vbOKOnly, "EzBiz"
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            tXTMEDICINE.SetFocus
    End Select
End Sub

Private Sub Form_Load()
    On Error GoTo eRRhAND
    
    db.Execute "Update itemmast set ITEM_SPEC = '' where isnull(ITEM_SPEC) "
    frmloadflag = True
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        If RSTCOMPANY!hide_ws = "Y" Then
            chkws.Value = 1
        Else
            chkws.Value = 0
        End If
        If RSTCOMPANY!hide_van = "Y" Then
            chkvp.Value = 1
        Else
            chkvp.Value = 0
        End If
        If RSTCOMPANY!hide_lwp = "Y" Then
            chklwp.Value = 1
        Else
            chklwp.Value = 0
        End If
        If RSTCOMPANY!hide_category = "Y" Then
            chkhidecat.Value = 1
        Else
            chkhidecat.Value = 0
        End If
        If RSTCOMPANY!hide_company = "Y" Then
            chkhidecomp.Value = 1
        Else
            chkhidecomp.Value = 0
        End If
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
           
    frmloadflag = False
    REPFLAG = True
    PHY_FLAG = True
    
'    If chkws.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 8) = ""
'        GRDSTOCK.ColWidth(8) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 8) = "WS"
'        GRDSTOCK.ColWidth(8) = 1000
'    End If
'    If chkvp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 9) = ""
'        GRDSTOCK.ColWidth(9) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 9) = "VP"
'        GRDSTOCK.ColWidth(9) = 1000
'    End If
'    If chklwp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 17) = ""
'        GRDSTOCK.ColWidth(17) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 17) = "L.W.Price"
'        GRDSTOCK.ColWidth(17) = 1000
'    End If
'    If chkhidecat.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 18) = ""
'        GRDSTOCK.ColWidth(18) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 18) = "Category"
'        GRDSTOCK.ColWidth(18) = 1000
'    End If
'    If chkhidecomp.value = 1 Then
'        GRDSTOCK.TextMatrix(0, 19) = ""
'        GRDSTOCK.ColWidth(19) = 0
'    Else
'        GRDSTOCK.TextMatrix(0, 19) = "Company"
'        GRDSTOCK.ColWidth(19) = 1000
'    End If
'    If (frmLogin.rs!Level = "0" or frmLogin.rs!Level = "4") Then
'        GRDSTOCK.TextMatrix(0, 11) = "Per Rate"
'        GRDSTOCK.TextMatrix(0, 12) = "Net Cost"
'        GRDSTOCK.TextMatrix(0, 20) = "Profit%"
'        GRDSTOCK.ColWidth(11) = 1000
'        GRDSTOCK.ColWidth(12) = 900
'        GRDSTOCK.ColWidth(20) = 1000
'    Else
'        GRDSTOCK.TextMatrix(0, 11) = ""
'        GRDSTOCK.TextMatrix(0, 12) = ""
'        GRDSTOCK.TextMatrix(0, 20) = ""
'        GRDSTOCK.ColWidth(11) = 0
'        GRDSTOCK.ColWidth(12) = 0
'        GRDSTOCK.ColWidth(20) = 0
'    End If
    Call Fillgrid
    DTFROM.Value = Format(Date, "DD/MM/YYYY")
    DTFROM.Value = Null
    Call Fillgrid
    'Me.Height = 8415
    'Me.Width = 6465
    Me.Left = 0
    Me.Top = 0
    Exit Sub
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    If RSTREP.State = 1 Then RSTREP.Close
    If PHY_REC.State = 1 Then PHY_REC.Close
    If adoGrid.State = 1 Then
        adoGrid.Close
        Set adoGrid = Nothing
    End If
    MDIMAIN.PCTMENU.Enabled = True
    MDIMAIN.PCTMENU.SetFocus
End Sub

Private Sub grdmsc_AfterColEdit(ByVal ColIndex As Integer)
'    Dim rststock, RSTITEMMAST As ADODB.Recordset
'
'    On Error GoTo eRRHAND
'
'        Select Case grdmsc.Col

'

'            Case 8  'WS
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_WS = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 13) = 1
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_WS = Val(TXTsample.Text)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 15  'CRTN_PACK
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CRTN_PACK = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Val(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!CRTN_PACK = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 16  'L. R. PRICE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_CRTN from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_CRTN) AND P_CRTN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_CRTN = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If grdmsc.TextMatrix(grdmsc.Row, 15) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_CRTN = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 17  'L. W. PRICE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_LWS from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_LWS) AND P_LWS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_LWS = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If grdmsc.TextMatrix(grdmsc.Row, 15) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_LWS = Val(TXTsample.Text)
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 9  'VAN
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_VAN = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 15) = 1
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 13) = 1
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
'                Do Until rststock.EOF
'                    rststock!P_WS = Val(TXTsample.Text)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 15)) = 0 Then
'                        rststock!CRTN_PACK = 1
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 13)) = 0 Then
'                        rststock!LOOSE_PACK = 1
'                    End If
'                    rststock.Update
'                    rststock.MoveNext
'                Loop
'                rststock.Close
'                Set rststock = Nothing
'
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 18  'CATEGORY
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!Category = Trim(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
''                Case 11  'LOC
''                    Set rststock = New ADODB.Recordset
''                    rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
''                    If Not (rststock.EOF And rststock.BOF) Then
''                        rststock!BIN_LOCATION = Trim(TXTsample.Text)
''                        grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
''                        rststock.Update
''                    End If
''                    rststock.Close
''                    Set rststock = Nothing
''                    grdmsc.Enabled = True
''                    TXTsample.Visible = False
''                    grdmsc.SetFocus
'
'            Case 11  'COST
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!ITEM_COST = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 12)) <> 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 7)) - Val(grdmsc.TextMatrix(grdmsc.Row, 12))) * 100) / Val(grdmsc.TextMatrix(grdmsc.Row, 12)), 2), "0.00")
'                    Else
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = 0
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 25) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 3))
'                Call Toatal_value
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 25  'VALUE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 3)) <> 0 Then
'                        rststock!ITEM_COST = Round(Val(TXTsample.Text) / Val(grdmsc.TextMatrix(grdmsc.Row, 3)), 3)
'                        grdmsc.TextMatrix(grdmsc.Row, 11) = Format(Round(Val(TXTsample.Text) / Val(grdmsc.TextMatrix(grdmsc.Row, 3)), 3), "0.000")
'                        grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                    End If
'                    If Val(grdmsc.TextMatrix(grdmsc.Row, 11)) <> 0 Then
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 7)) - Val(grdmsc.TextMatrix(grdmsc.Row, 11))) * 100) / Val(grdmsc.TextMatrix(grdmsc.Row, 11)), 2), "0.00")
'                    Else
'                        grdmsc.TextMatrix(grdmsc.Row, 20) = 0
'                    End If
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 25) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 3))
'                Call Toatal_value
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 6  'MRP
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If rststock.RecordCount > 1 Then
'                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
'                        rststock.Close
'                        Set rststock = Nothing
'                        TXTsample.SetFocus
'                        Exit Sub
'                    End If
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!MRP = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'
'            Case 10  'TAX
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!SALES_TAX = Val(TXTsample.Text)
'                    rststock!CHECK_FLAG = "V"
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Val(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.TextMatrix(grdmsc.Row, 12) = Val(grdmsc.TextMatrix(grdmsc.Row, 11)) + (Val(grdmsc.TextMatrix(grdmsc.Row, 11)) * Val(grdmsc.TextMatrix(grdmsc.Row, 10)) / 100)
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'                Call Toatal_value
'            Case 20  'Profit %
'                grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.000")
'                grdmsc.TextMatrix(grdmsc.Row, 7) = Format(Round(((Val(grdmsc.TextMatrix(grdmsc.Row, 12)) * grdmsc.TextMatrix(grdmsc.Row, 13)) * Val(TXTsample.Text) / 100) + (Val(grdmsc.TextMatrix(grdmsc.Row, 12)) * grdmsc.TextMatrix(grdmsc.Row, 13)), 2), "0.000")
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!P_RETAIL = Val(grdmsc.TextMatrix(grdmsc.Row, 7))
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 21  'Cust Disc
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CUST_DISC = Val(TXTsample.Text)
'                    rststock!DISC_AMT = 0
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 22) = ""
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 22  'Cust Disc Amt
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!DISC_AMT = Val(TXTsample.Text)
'                    rststock!CUST_DISC = 0
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    grdmsc.TextMatrix(grdmsc.Row, 21) = ""
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 14  'HSN CODE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!REMARKS = Trim(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Trim(TXTsample.Text)
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 26  'CESS%
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CESS_PER = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 27  'CESS RATE
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!CESS_AMT = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
'            Case 13  'UNIT
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT * from ITEMMAST where ITEMMAST.ITEM_CODE = '" & grdmsc.TextMatrix(grdmsc.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    rststock!LOOSE_PACK = Val(TXTsample.Text)
'                    grdmsc.TextMatrix(grdmsc.Row, grdmsc.Col) = Format(Val(TXTsample.Text), "0.00")
'                    rststock.Update
'                End If
'                rststock.Close
'                Set rststock = Nothing
'                grdmsc.Enabled = True
'                TXTsample.Visible = False
'                grdmsc.SetFocus
                
'        End Select
'        Exit Sub
'eRRHAND:
'    MsgBox Err.Description
End Sub

Private Sub grdmsc_AfterColUpdate(ByVal ColIndex As Integer)
    Dim rststock As ADODB.Recordset
    
    On Error GoTo eRRhAND
    
    Select Case ColIndex
        Case 0  ' Item Code
            db.Execute "Update RTRXFILE set ITEM_CODE = '" & grdmsc.Columns(0) & "' where ITEM_CODE = '" & grdmsc.Tag & "' "
            db.Execute "Update TRXFILE set ITEM_CODE = '" & grdmsc.Columns(0) & "' where ITEM_CODE = '" & grdmsc.Tag & "' "
        Case 1  ' Item Name
            db.Execute "Update RTRXFILE set ITEM_NAME = '" & grdmsc.Columns(1) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' "
            db.Execute "Update TRXFILE set ITEM_NAME = '" & grdmsc.Columns(1) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' "
        Case 4
                Dim INWARD, OUTWARD, BAL_QTY As Double
                Dim TRXMAST As ADODB.Recordset
                Dim RSTITEMMAST As ADODB.Recordset

                Screen.MousePointer = vbHourglass
                Set RSTITEMMAST = New ADODB.Recordset
                RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & grdmsc.Columns(0) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
                If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
                    INWARD = 0
                    OUTWARD = 0
                    BAL_QTY = 0
                    
                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        INWARD = INWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        INWARD = INWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        BAL_QTY = BAL_QTY + IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) '* IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing

                    Set rststock = New ADODB.Recordset
                    rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='MI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly
                    Do Until rststock.EOF
                        OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                        OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
                        rststock.MoveNext
                    Loop
                    rststock.Close
                    Set rststock = Nothing
                    
                    Dim BILL_NO, M_DATA As Double
                    Set TRXMAST = New ADODB.Recordset
                    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'ST'", db, adOpenStatic, adLockReadOnly
                    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
                        BILL_NO = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
                    End If
                    TRXMAST.Close
                    Set TRXMAST = Nothing

                    If Not (grdmsc.Columns(2) - (Val(INWARD - OUTWARD)) = 0) Then
                        Set rststock = New ADODB.Recordset
                        rststock.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'ST' ", db, adOpenStatic, adLockOptimistic, adCmdText
                        'If (rststock.EOF And rststock.BOF) Then
                            rststock.AddNew
                            rststock!TRX_TYPE = "ST"
                            rststock!TRX_YEAR = Year(MDIMAIN.DTFROM.Value)
                            rststock!VCH_NO = BILL_NO
                            rststock!LINE_NO = 1
                            rststock!ITEM_CODE = RSTITEMMAST!ITEM_CODE
                        'End If
                        rststock!BAL_QTY = grdmsc.Columns(2) - (Val(BAL_QTY))
                        rststock!QTY = grdmsc.Columns(2) - (Val(INWARD - OUTWARD))
                        rststock!TRX_TOTAL = 0
                        rststock!VCH_DATE = Format(DTFROM.Value, "dd/mm/yyyy")
                        rststock!ITEM_NAME = grdmsc.Columns(1)
                        rststock!item_COST = grdmsc.Columns(9)
                        rststock!LINE_DISC = 1
                        rststock!P_DISC = 0
                        rststock!MRP = grdmsc.Columns(4)
                        rststock!PTR = grdmsc.Columns(9)
                        rststock!SALES_PRICE = grdmsc.Columns(5)
                        rststock!P_RETAIL = Val(grdmsc.Columns(5))
                        rststock!P_WS = Val(grdmsc.Columns(6))
                        rststock!P_VAN = Val(grdmsc.Columns(7))
                        rststock!P_CRTN = Val(grdmsc.Columns(10))
                        rststock!P_LWS = Val(grdmsc.Columns(11))
                        rststock!CRTN_PACK = IIf(IsNull(RSTITEMMAST!CRTN_PACK) Or RSTITEMMAST!CRTN_PACK = 0, 1, RSTITEMMAST!CRTN_PACK)
                        rststock!Category = grdmsc.Columns(12)
                        rststock!gross_amt = 0
                        rststock!COM_FLAG = "P"
                        rststock!COM_PER = 0
                        rststock!COM_AMT = 0
                        rststock!SALES_TAX = Val(grdmsc.Columns(8))
                        rststock!LOOSE_PACK = RSTITEMMAST!LOOSE_PACK
                        rststock!PACK_TYPE = grdmsc.Columns(3)
                        rststock!WARRANTY = Null
                        rststock!WARRANTY_TYPE = Null
                        rststock!UNIT = 1 'Val(grdsales.TextMatrix(Val(TXTSLNO.Text), 4))
                        'rststock!VCH_DESC = "Received From " & DataList2.Text
                        rststock!REF_NO = ""
                        'rststock!ISSUE_QTY = 0
                        rststock!CST = 0
                        rststock!DISC_FLAG = "P"
                        rststock!SCHEME = 0
                        'rststock!EXP_DATE = Null
                        rststock!FREE_QTY = 0
                        rststock!CREATE_DATE = Format(Date, "dd/mm/yyyy")
                        rststock!C_USER_ID = "SM"
                        rststock!check_flag = "V"

                        'rststock!M_USER_ID = DataList2.BoundText
                        'rststock!PINV = Trim(TXTINVOICE.Text)
                        rststock.Update
                        rststock.Close
                        Set rststock = Nothing

                        'RSTITEMMAST!CLOSE_QTY = grdmsc.Columns(2)
                        'RSTITEMMAST!RCPT_QTY = INWARD + grdmsc.Columns(2)
                        'RSTITEMMAST!ISSUE_QTY = OUTWARD
                        'RSTITEMMAST.Update
                    End If
                    RSTITEMMAST.Close
                    Set RSTITEMMAST = Nothing
                End If
                Screen.MousePointer = vbNormal
            Case 6
                db.Execute "Update RTRXFILE set P_RETAIL = '" & grdmsc.Columns(5) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
            Case 7
                db.Execute "Update RTRXFILE set P_WS = '" & grdmsc.Columns(5) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
            Case 8
                db.Execute "Update RTRXFILE set P_VAN = '" & grdmsc.Columns(7) & "' where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 "
                
                

               

            
    End Select
    Exit Sub
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub grdmsc_BeforeColUpdate(ByVal ColIndex As Integer, OldValue As Variant, Cancel As Integer)
    Dim rststock As ADODB.Recordset
    On Error GoTo eRRhAND
    Select Case ColIndex
        Case 0
            grdmsc.Tag = OldValue
        Case 4
            If IsNull(DTFROM.Value) Then
                MsgBox "Select the Date for Opening Qty", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
            If (DTFROM.Value) > Date Then
                MsgBox "The date could not be greater than Today", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
            If Val(grdmsc.Columns(9)) = 0 Then
                MsgBox "Please enter the cost", vbOKOnly, "Price Analysis"
                Cancel = 1
                Exit Sub
            End If
        
        Case 6
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT P_RETAIL from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_RETAIL) AND P_RETAIL <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing

            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing

            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
            If rststock.RecordCount > 1 Then
                If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                    rststock.Close
                    Set rststock = Nothing
                    Cancel = 1
                    Exit Sub
                End If
            End If
            rststock.Close
            Set rststock = Nothing
    
        Case 7
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT P_WS from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_WS) AND P_WS <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
        
        Case 8
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT P_VAN from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(P_VAN) AND P_VAN <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("The Price will be affected to all the existing qty. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT MRP from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(MRP) AND MRP <>0 ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different MRPs Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    
                Set rststock = New ADODB.Recordset
                rststock.Open "SELECT DISTINCT REF_NO from RTRXFILE where ITEM_CODE = '" & grdmsc.Columns(0) & "' AND BAL_QTY >0 AND NOT ISNULL(REF_NO) AND REF_NO <> '' ", db, adOpenStatic, adLockReadOnly, adCmdText
                If rststock.RecordCount > 1 Then
                    If MsgBox("Different Batches Available. Are You Sure?", vbYesNo + vbDefaultButton2, "Price Change") = vbNo Then
                        rststock.Close
                        Set rststock = Nothing
                        Cancel = 1
                        Exit Sub
                    End If
                End If
                rststock.Close
                Set rststock = Nothing
    End Select
    Exit Sub
eRRhAND:
    Cancel = 1
    MsgBox err.Description
    
End Sub

Private Sub OptAll_Click()
    tXTMEDICINE.SetFocus
End Sub

Private Sub TxtHSNCODE_Change()
    Call tXTMEDICINE_Change
End Sub

Private Sub TXTITEMCODE_Change()
    Call tXTMEDICINE_Change
End Sub

Private Sub tXTMEDICINE_Change()
    On Error GoTo eRRhAND
    Call Fillgrid
    If REPFLAG = True Then
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If CHKCATEGORY2.Value = 0 Then
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        Else
            If OptStock.Value = True Then
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
            Else
                RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF'  AND ucase(CATEGORY) <> 'OWN' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND MANUFACTURER = '" & DataList1.BoundText & "' AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%' AND (ISNULL(ITEM_SPEC) OR ITEM_SPEC Like '%" & Me.txtspecs.Text & "%') ORDER BY ITEM_NAME", db, adOpenForwardOnly
            End If
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If DataList2.VisibleCount = 0 Then Exit Sub
            TxtCode.SetFocus
        Case vbKeyEscape
            Call CmdExit_Click
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


Private Sub TXTCODE_Change()
    Call tXTMEDICINE_Change
    Exit Sub
    On Error GoTo eRRhAND
    Call Fillgrid
    If REPFLAG = True Then
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    Else
        RSTREP.Close
        If OptStock.Value = True Then
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenForwardOnly
        Else
            RSTREP.Open "Select DISTINCT ITEM_CODE, ITEM_NAME From ITEMMAST  WHERE  ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' ORDER BY ITEM_NAME", db, adOpenForwardOnly
        End If
        REPFLAG = False
    End If
    Set Me.DataList2.RowSource = RSTREP
    DataList2.ListField = "ITEM_NAME"
    DataList2.BoundColumn = "ITEM_CODE"
    
    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRhAND:
    MsgBox err.Description
End Sub

Private Sub TxtCode_GotFocus()
    TxtCode.SelStart = 0
    TxtCode.SelLength = Len(TxtCode.Text)
    'Call Fillgrid
End Sub

Private Sub TxtCode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TXTITEMCODE.SetFocus
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


Private Function Toatal_value()
    Dim Stk_Val As Double
    Dim i As Long
    lblpvalue.Caption = ""
    lblnetvalue.Caption = ""
'    For i = 1 To GRDSTOCK.Rows - 1
'        lblpvalue.Caption = Format(Val(lblpvalue.Caption) + Val(GRDSTOCK.TextMatrix(i, 25)), "0.00")
''        If (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))) <> Val(GRDSTOCK.TextMatrix(i, 25)) Then
''            MsgBox ""
''        End If
'        lblnetvalue.Caption = Format(Val(lblnetvalue.Caption) + (Val(GRDSTOCK.TextMatrix(i, 12)) * Val(GRDSTOCK.TextMatrix(i, 3))), "0.00")
'    Next i
End Function


Private Sub TXTDEALER2_Change()
    
    On Error GoTo eRRhAND
    If flagchange2.Caption <> "1" Then
        If chkcategory.Value = 1 Then
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT CATEGORY From CATEGORY WHERE CATEGORY Like '" & TXTDEALER2.Text & "%' ORDER BY CATEGORY", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!Category
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "CATEGORY"
            DataList1.BoundColumn = "CATEGORY"
        Else
            If PHY_FLAG = True Then
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            Else
                PHY_REC.Close
                PHY_REC.Open "Select DISTINCT MANUFACTURER From MANUFACT WHERE MANUFACTURER Like '" & TXTDEALER2.Text & "%' ORDER BY MANUFACTURER", db, adOpenStatic, adLockReadOnly
                PHY_FLAG = False
            End If
            If (PHY_REC.EOF And PHY_REC.BOF) Then
                LBLDEALER2.Caption = ""
            Else
                LBLDEALER2.Caption = PHY_REC!MANUFACTURER
            End If
            Set Me.DataList1.RowSource = PHY_REC
            DataList1.ListField = "MANUFACTURER"
            DataList1.BoundColumn = "MANUFACTURER"

        End If
    End If
    Exit Sub
eRRhAND:
    MsgBox err.Description
    
End Sub


Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
    'CHKCATEGORY2.value = 1
End Sub

Private Sub TXTDEALER2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If DataList1.VisibleCount = 0 Then Exit Sub
            'lbladdress.Caption = ""
            DataList1.SetFocus
    End Select

End Sub

Private Sub TXTDEALER2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_Click()
        
    TXTDEALER2.Text = DataList1.Text
    LBLDEALER2.Caption = TXTDEALER2.Text
    Call Fillgrid
    tXTMEDICINE.SetFocus
End Sub

Private Sub DataList1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTDEALER2.Text) = "" Then Exit Sub
            If IsNull(DataList1.SelectedItem) Then
                MsgBox "Select Category From List", vbOKOnly, "Category List..."
                DataList1.SetFocus
                Exit Sub
            End If
        Case vbKeyEscape
            TXTDEALER2.SetFocus
    End Select
End Sub

Private Sub DataList1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" "), Asc("("), Asc(")")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList1_GotFocus()
    flagchange2.Caption = 1
    TXTDEALER2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTDEALER2.Text
    Call DataList1_Click
    'CHKCATEGORY2.value = 1
End Sub

Private Sub DataList1_LostFocus()
     flagchange2.Caption = ""
End Sub

Private Sub TxtItemcode_GotFocus()
    TXTITEMCODE.SelStart = 0
    TXTITEMCODE.SelLength = Len(TXTITEMCODE.Text)
    'Call Fillgrid
End Sub

Private Sub TxtItemcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtHSNCODE.SetFocus
        Case vbKeyEscape
            TxtCode.SetFocus
    End Select

End Sub

Private Sub TxtItemcode_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Sub TxtHSNCODE_GotFocus()
    TxtHSNCODE.SelStart = 0
    TxtHSNCODE.SelLength = Len(TxtHSNCODE.Text)
    'Call Fillgrid
End Sub

Private Sub TxtHSNCODE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'Call CmdLoad_Click
        Case vbKeyEscape
            TXTITEMCODE.SetFocus
    End Select

End Sub

Private Sub TxtHSNCODE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case Else
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End Select
End Sub

Private Function Fillgrid()
    On Error GoTo eRRhAND
    Screen.MousePointer = vbHourglass
    Set grdmsc.DataSource = Nothing
    Set adoGrid = New ADODB.Recordset
    With adoGrid
        .CursorLocation = adUseClient
        If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                ElseIf OptPC.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    ElseIf OptPC.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    End If
                Else
                    If OptStock.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    ElseIf OptPC.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
                    End If
        
                End If
            End If
        Else
            If CHKCATEGORY2.Value = 0 And chkcategory.Value = 0 Then
                If OptStock.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF') AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                ElseIf OptPC.Value = True Then
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                Else
                    .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%')) AND ucase(CATEGORY) <> 'SERVICE CHARGE' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                End If
            Else
                If CHKCATEGORY2.Value = 1 Then
                    If OptStock.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    ElseIf OptPC.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE (ISNULL(CATEGORY) OR ucase(CATEGORY) <> 'SELF')  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    End If
                Else
                    If OptStock.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' AND CLOSE_QTY <>0 ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    ElseIf OptPC.Value = True Then
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE PRICE_CHANGE = 'Y' and ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    Else
                        .Open "SELECT ITEM_CODE, ITEM_NAME, BIN_LOCATION, ITEM_MAL, CLOSE_QTY, MRP, P_RETAIL, P_WS, P_VAN, SALES_TAX, ITEM_SPEC FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF'  AND ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtCode.Text & "%' AND ITEM_CODE Like '%" & Me.TXTITEMCODE.Text & "%' AND ITEM_SPEC Like '%" & Me.txtspecs.Text & "%' AND (ISNULL(BIN_LOCATION) OR BIN_LOCATION Like '%" & Me.TxtHSNCODE.Text & "%') AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                    End If
        
                End If
            End If
        End If
    End With
    Set grdmsc.DataSource = adoGrid
    
    grdmsc.Columns(0).Caption = "ITEM CODE"
    grdmsc.Columns(1).Caption = "ITEM NAME"
    grdmsc.Columns(2).Caption = "LOCATION"
    grdmsc.Columns(3).Caption = "MALAYALAM"
    grdmsc.Columns(4).Caption = "QTY"
    grdmsc.Columns(5).Caption = "MRP"
    grdmsc.Columns(6).Caption = "R. PRICE"
    grdmsc.Columns(7).Caption = "W. PRICE"
    grdmsc.Columns(8).Caption = "V. PRICE"
    grdmsc.Columns(9).Caption = "TAX %"
    grdmsc.Columns(10).Caption = "SPECIFICATIONS"
        
    grdmsc.Columns(0).Width = 1100
    grdmsc.Columns(1).Width = 5000
    grdmsc.Columns(2).Width = 1400
    grdmsc.Columns(3).Width = 2500
    grdmsc.Columns(4).Width = 900
    grdmsc.Columns(5).Width = 900
    grdmsc.Columns(6).Width = 900
    grdmsc.Columns(7).Width = 900
    grdmsc.Columns(8).Width = 900
    grdmsc.Columns(9).Width = 900
    grdmsc.Columns(10).Width = 2000
    
    Screen.MousePointer = vbNormal
    Exit Function
eRRhAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Function

Private Sub txtspecs_Change()
    Call tXTMEDICINE_Change
End Sub

Private Sub txtspecs_GotFocus()
    txtspecs.SelStart = 0
    txtspecs.SelLength = Len(txtspecs.Text)
End Sub

Private Sub txtspecs_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call Fillgrid
        Case vbKeyEscape
            TxtHSNCODE.SetFocus
    End Select
End Sub

Private Sub txtspecs_KeyPress(KeyAscii As Integer)
     Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub
