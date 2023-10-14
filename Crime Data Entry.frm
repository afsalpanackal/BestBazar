VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmCrimedata 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PURCHASE ORDER"
   ClientHeight    =   9675
   ClientLeft      =   1095
   ClientTop       =   405
   ClientWidth     =   15105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "Crime Data Entry.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   MouseIcon       =   "Crime Data Entry.frx":000C
   ScaleHeight     =   170.656
   ScaleMode       =   6  'Millimeter
   ScaleWidth      =   266.436
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton CMDDELITEM 
      Caption         =   "&Delete Item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1380
      TabIndex        =   15
      Top             =   7995
      Width           =   1260
   End
   Begin VB.CommandButton cmdstockcorrect 
      Caption         =   "&Stock Correction"
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
      Left            =   12135
      TabIndex        =   21
      Top             =   9075
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSFlexGridLib.MSFlexGrid grdSTOCKLESS 
      Height          =   8625
      Left            =   10425
      TabIndex        =   18
      Top             =   15
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   15214
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   0
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
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Height          =   405
      Left            =   -255
      TabIndex        =   31
      Top             =   -495
      Width           =   11850
      Begin VB.Label LBLSHOP 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MEDICALS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   450
         TabIndex        =   32
         Top             =   120
         Width           =   11010
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF00&
      Height          =   9900
      Left            =   45
      TabIndex        =   22
      Top             =   -270
      Width           =   15090
      Begin VB.Frame Frame5 
         BackColor       =   &H00FFFF80&
         Height          =   735
         Left            =   45
         TabIndex        =   42
         Top             =   8700
         Width           =   3945
         Begin VB.Label LBLMANUFACT 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "COMPANY"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   120
            TabIndex        =   43
            Top             =   240
            Width           =   3705
            WordWrap        =   -1  'True
         End
      End
      Begin VB.CheckBox CHKSELECT 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF00&
         Caption         =   "Select All"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   8820
         TabIndex        =   41
         Top             =   7680
         Width           =   1380
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFF00&
         Height          =   3690
         Left            =   4020
         TabIndex        =   33
         Top             =   6180
         Width           =   4785
         Begin VB.ListBox lstselDist 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   3090
            Left            =   45
            Style           =   1  'Checkbox
            TabIndex        =   13
            Top             =   120
            Width           =   4665
         End
      End
      Begin VB.TextBox TXTsample 
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
         Height          =   290
         Left            =   3000
         TabIndex        =   40
         Top             =   3105
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSDataListLib.DataCombo CMBSUPPLIER 
         Height          =   330
         Left            =   2100
         TabIndex        =   39
         Top             =   3630
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   255
         Text            =   ""
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
      Begin VB.CommandButton CMDORDER 
         Caption         =   "ADD STOCK &LESS ITEMS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   37
         Top             =   8760
         Visible         =   0   'False
         Width           =   1470
      End
      Begin VB.CommandButton CMDNINQTY 
         Caption         =   "SET &MIN QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   2700
         TabIndex        =   16
         Top             =   8265
         Width           =   1260
      End
      Begin VB.ComboBox CmbQty 
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
         ItemData        =   "Crime Data Entry.frx":0316
         Left            =   4710
         List            =   "Crime Data Entry.frx":0323
         TabIndex        =   4
         Top             =   1380
         Width           =   1095
      End
      Begin VB.CommandButton CMDVALUE 
         Caption         =   "APPROX VALUE"
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
         Height          =   495
         Left            =   8850
         TabIndex        =   17
         Top             =   8940
         Width           =   1500
      End
      Begin VB.CommandButton cmditemcreate 
         Caption         =   "&Create Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   60
         TabIndex        =   14
         Top             =   8265
         Width           =   1230
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
         Height          =   420
         Left            =   1590
         TabIndex        =   10
         Top             =   6795
         Width           =   1155
      End
      Begin VB.CommandButton CMDDELETESTOCK 
         BackColor       =   &H00400000&
         Caption         =   "DELETE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   10395
         MaskColor       =   &H80000007&
         TabIndex        =   19
         Top             =   8940
         UseMaskColor    =   -1  'True
         Width           =   1155
      End
      Begin VB.TextBox txtstockcorrect 
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
         Left            =   11985
         MaxLength       =   25
         TabIndex        =   20
         Top             =   9090
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.ListBox lstmanufact 
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
         Height          =   5880
         Left            =   7020
         Style           =   1  'Checkbox
         TabIndex        =   34
         Top             =   285
         Width           =   3330
      End
      Begin VB.CommandButton cmdadjust 
         BackColor       =   &H00400000&
         Caption         =   "Stock Adjust"
         Height          =   345
         Left            =   5835
         MaskColor       =   &H80000007&
         TabIndex        =   5
         Top             =   1365
         UseMaskColor    =   -1  'True
         Width           =   1170
      End
      Begin VB.CommandButton CmdAdd 
         BackColor       =   &H00400000&
         Caption         =   "&ADD"
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
         Left            =   90
         MaskColor       =   &H80000007&
         TabIndex        =   6
         Top             =   6345
         UseMaskColor    =   -1  'True
         Width           =   1185
      End
      Begin VB.TextBox TxtQty 
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
         Height          =   285
         Left            =   4065
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1395
         Width           =   600
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
         Left            =   75
         TabIndex        =   0
         Top             =   300
         Width           =   3645
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1035
         Left            =   75
         TabIndex        =   1
         Top             =   660
         Width           =   3645
         _ExtentX        =   6429
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
      Begin MSDataListLib.DataList LSTDISTI 
         Height          =   1035
         Left            =   3765
         TabIndex        =   2
         Top             =   300
         Width           =   3225
         _ExtentX        =   5689
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
      Begin VB.ListBox LSTDUMMY 
         Height          =   1425
         Left            =   5685
         TabIndex        =   30
         Top             =   6585
         Visible         =   0   'False
         Width           =   2580
      End
      Begin VB.Frame frmecontrol 
         BackColor       =   &H00FFFF00&
         Height          =   1125
         Left            =   30
         TabIndex        =   24
         Top             =   6165
         Width           =   3990
         Begin VB.CommandButton CMDCLEAR 
            Caption         =   "Clear All &Items"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   60
            TabIndex        =   9
            Top             =   630
            Width           =   1410
         End
         Begin VB.CommandButton CmdDelete 
            Caption         =   "DE&LETE"
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
            Left            =   2685
            TabIndex        =   7
            Top             =   180
            Width           =   1170
         End
         Begin VB.CommandButton Create 
            Caption         =   "&PRINT"
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
            TabIndex        =   8
            Top             =   180
            Width           =   1170
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFF00&
         Height          =   1350
         Left            =   8850
         TabIndex        =   36
         Top             =   6180
         Width           =   1500
         Begin VB.CommandButton CMDADDDIST 
            BackColor       =   &H00400000&
            Caption         =   "Add Distributor"
            Height          =   495
            Left            =   90
            MaskColor       =   &H80000007&
            TabIndex        =   11
            Top             =   165
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
         Begin VB.CommandButton CMDREMOVE 
            BackColor       =   &H00400000&
            Caption         =   "Remove Distributor"
            Height          =   495
            Left            =   90
            MaskColor       =   &H80000007&
            TabIndex        =   12
            Top             =   720
            UseMaskColor    =   -1  'True
            Width           =   1275
         End
      End
      Begin MSFlexGridLib.MSFlexGrid grdorder 
         Height          =   4425
         Left            =   45
         TabIndex        =   38
         Top             =   1740
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   7805
         _Version        =   393216
         Rows            =   1
         Cols            =   8
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
      Begin VB.Label LBLVALUE 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   8820
         TabIndex        =   35
         Top             =   8475
         Width           =   1530
      End
      Begin VB.Label LBLMINI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "MINI STOCK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   2460
         TabIndex        =   29
         Top             =   7740
         Width           =   885
      End
      Begin VB.Label LBLMN 
         BackStyle       =   0  'Transparent
         Caption         =   "MINIMUM STOCK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   105
         TabIndex        =   28
         Top             =   7725
         Width           =   2160
      End
      Begin VB.Label LBLAVSTOCK 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "AVL STOCK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   2490
         TabIndex        =   27
         Top             =   7380
         Width           =   840
      End
      Begin VB.Label LBLAV 
         BackStyle       =   0  'Transparent
         Caption         =   "AVAILABLE STOCK"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   375
         Left            =   105
         TabIndex        =   26
         Top             =   7365
         Width           =   2355
      End
      Begin VB.Label LBLUNIT 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Left            =   6690
         TabIndex        =   25
         Top             =   75
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   3750
         TabIndex        =   23
         Top             =   1410
         Width           =   405
      End
   End
End
Attribute VB_Name = "FrmCrimedata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY As New ADODB.Recordset
Dim RSTREP As New ADODB.Recordset
Dim rstTMP As New ADODB.Recordset

Dim TMPFLAG As Boolean 'TMP
Dim REPFLAG As Boolean 'REP
Dim SEARCHFLAG As Boolean 'SEARCH BY NAME
Dim M_EDIT As Boolean
Dim RCVDFLAG As Boolean

Dim k As Integer
Dim CLOSEALL As Integer
  
Private Sub CHKSELECT_Click()
    Dim i As Integer
    If CHKSELECT.Value = 1 Then
        For i = 0 To lstselDist.ListCount - 1
            lstselDist.Selected(i) = True
        Next i
        lstselDist.Refresh
        lstselDist.SetFocus
    Else
        For i = 0 To lstselDist.ListCount - 1
            lstselDist.Selected(i) = False
        Next i
        lstselDist.Refresh
        lstselDist.SetFocus
    End If
End Sub

Private Sub CmbQty_KEyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.SetFocus
    End Select
End Sub

Private Sub CMBSUPPLIER_GotFocus()
    CMBSUPPLIER.SelStart = 0
    CMBSUPPLIER.SelLength = Len(CMBSUPPLIER.Text)
End Sub

Private Sub CMBSUPPLIER_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTSUPPLIER As ADODB.Recordset
    
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(CMBSUPPLIER.Text) = "" Then Exit Sub
            Set RSTSUPPLIER = New ADODB.Recordset
            RSTSUPPLIER.Open "SELECT ACT_NAME, ACT_CODE from [TMPORDERLIST] Where ITEM_CODE = '" & grdorder.TextMatrix(grdorder.Row, 1) & "' and ACT_CODE = '" & grdorder.TextMatrix(grdorder.Row, 5) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
            If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                RSTSUPPLIER!ACT_NAME = Trim(CMBSUPPLIER.Text)
                RSTSUPPLIER!ACT_code = CMBSUPPLIER.BoundText
                RSTSUPPLIER.Update
            End If
            RSTSUPPLIER.Close
            Set RSTSUPPLIER = Nothing
            
            grdorder.TextMatrix(grdorder.Row, 4) = Trim(CMBSUPPLIER.Text)
            grdorder.TextMatrix(grdorder.Row, 5) = CMBSUPPLIER.BoundText
            grdorder.Enabled = True
            CMBSUPPLIER.Visible = False
            grdorder.SetFocus
            Call fillSUPPLIERLIST
        Case vbKeyEscape
            CMBSUPPLIER.Visible = False
            grdorder.SetFocus
    End Select
End Sub

Private Sub CMBSUPPLIER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMBSUPPLIER_LostFocus()
    TXTsample.Visible = False
    CMBSUPPLIER.Visible = False
End Sub

Private Sub CMDADD_Click()
    Dim rstpurchase As ADODB.Recordset
    Dim i As Integer
    Dim FCODE, FCODE2, FCODE3, FCODE4 As String
    Dim FCODE1 As Integer
    Dim FCODE6 As Integer
    'FCODE - ITEM  FCODE1 - QTY , FCODE2 - DISTI, FCODE3 - ITEM CODE, FCODE4 - DISTI CODE, FCODE5 - FLAG
    
    If DataList2.BoundText = "" Then
        MsgBox "SELECT THE ITEM", vbOKOnly, "ORDER"
        tXTMEDICINE.SetFocus
        Exit Sub
    End If
    
    Set rstpurchase = New ADODB.Recordset
    rstpurchase.Open "select * from [TmpOrderlist] Where ITEM_CODE = '" & DataList2.BoundText & "' and ACT_CODE = '" & LSTDISTI.BoundText & "'", db2, adOpenForwardOnly
    If rstpurchase.RecordCount > 0 Then
        MsgBox "Already Entered for " & rstpurchase!ACT_NAME, vbOKOnly, "PURCHASE ORDER"
        grdSTOCKLESS.SetFocus
        Exit Sub
    End If
    rstpurchase.Close
    Set rstpurchase = Nothing
    
    Set rstpurchase = New ADODB.Recordset
    rstpurchase.Open "select * from [TmpOrderlist] Where ITEM_CODE = '" & DataList2.BoundText & "'", db2, adOpenForwardOnly
    Do Until rstpurchase.EOF
        If (MsgBox("ALREADY ENTERED FOR " & rstpurchase!ACT_NAME & Chr(13) & Chr(13) & "Do You Want to Order it Again for " & LSTDISTI.Text & "?", vbYesNo, "PURCHASE ORDER") = vbNo) Then
            grdSTOCKLESS.SetFocus
            Exit Sub
        End If
        rstpurchase.MoveNext
    Loop
    rstpurchase.Close
    Set rstpurchase = Nothing
    
    If Trim(TXTQTY.Text) = "" Then
        MsgBox "Enter the Quantity", vbOKOnly, "ORDER"
        TXTQTY.SetFocus
        Exit Sub
    End If
    
    If LSTDISTI.Text = "" Then
        MsgBox "SELECT DISTRIBUTOR", vbOKOnly, "ORDER"
        LSTDISTI.SetFocus
        Exit Sub
    End If

    On Error GoTo eRRHAND
    FCODE = Me.DataList2.Text 'ITEM
    FCODE1 = Val(Me.TXTQTY.Text) 'QTY
    FCODE2 = Me.LSTDISTI.Text 'DISTRIBUTOR
    FCODE3 = Me.DataList2.BoundText
    FCODE4 = Me.LSTDISTI.BoundText
    FCODE5 = "N"  'FLAG
    FCODE6 = Val(LBLUNIT.Caption)
    
    db2.Execute ("insert into TmpOrderlist values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "')")
    
    
    If RCVDFLAG = True Then
        db2.Execute ("DELETE from [NONRCVD] WHERE NONRCVD.ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'")
        Call fillSTOCKLESSGRID
        RCVDFLAG = False
    End If
    
    Call fillORDERGRID
    Call fillSUPPLIERLIST
    
    TXTQTY.Text = ""
    CmbQty.Text = ""
    frmecontrol.Enabled = True
    grdorder.Enabled = True
    If M_EDIT = True Then
        tXTMEDICINE.Text = ""
        grdorder.Row = Val(grdorder.Tag)
        grdorder.SetFocus
    Else
        'grdorder.Row = grdorder.ApproxCount - 1
        tXTMEDICINE.SetFocus
    End If

    M_EDIT = False
    
   Exit Sub
   
eRRHAND:
    MsgBox "ALREADY ENTERED", vbOKOnly, "ORDER"
    
End Sub

Private Sub CMDADDDIST_Click()
    
    Dim RSTA As ADODB.Recordset
    Dim RSTB As ADODB.Recordset
    Dim RSTC As ADODB.Recordset
    
    Dim i As Integer
    
    If DataList2.BoundText = "" Then
        MsgBox "SELECT THE ITEM", vbOKOnly, "ORDER"
        tXTMEDICINE.SetFocus
        Exit Sub
    End If
    
    If lstmanufact.SelCount = 0 Then
        MsgBox "Please Select the Distributor to be added", vbOKOnly, "ORDER"
        Exit Sub
    End If

    
    On Error GoTo eRRHAND
    
    i = 0
    
    Set RSTA = New ADODB.Recordset
    
    RSTA.Open "SELECT *  FROM PRODLINK WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Set RSTB = New ADODB.Recordset
    RSTB.Open "SELECT *  FROM PRODLINK WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    With RSTA
        If Not (.EOF And .BOF) Then
            For i = 0 To lstmanufact.ListCount - 1
                If lstmanufact.Selected(i) Then
                    .AddNew
                    !ITEM_CODE = RSTB!ITEM_CODE
                    !ITEM_NAME = RSTB!ITEM_NAME
                    !RQTY = RSTB!RQTY
                    !ITEM_COST = RSTB!ITEM_COST
                    !MRP = RSTB!MRP
                    !PTR = RSTB!PTR
                    !SALES_PRICE = RSTB!SALES_PRICE
                    !SALES_TAX = RSTB!SALES_TAX
                    !UNIT = RSTB!UNIT
                    !Remarks = RSTB!Remarks
                    !ORD_QTY = RSTB!ORD_QTY
                    !CST = RSTB!CST
                    !ACT_code = Mid(lstmanufact.List(i), 1, 6)
                    !CREATE_DATE = RSTB!CREATE_DATE
                    !C_USER_ID = RSTB!C_USER_ID
                    !MODIFY_DATE = RSTB!MODIFY_DATE
                    !M_USER_ID = RSTB!M_USER_ID
                    !CHECK_FLAG = RSTB!CHECK_FLAG
                    !SITEM_CODE = RSTB!SITEM_CODE
                    .Update
                End If
            Next i
            
        Else
            Set RSTC = New ADODB.Recordset
            RSTC.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            For i = 0 To lstmanufact.ListCount - 1
                If lstmanufact.Selected(i) Then
                    .AddNew
                    !ITEM_CODE = RSTC!ITEM_CODE
                    !ITEM_NAME = RSTC!ITEM_NAME
                    !RQTY = Null
                    !ITEM_COST = RSTC!ITEM_COST
                    !MRP = RSTC!MRP
                    !PTR = RSTC!PTR
                    !SALES_PRICE = 0 'Val(RSTC!SALES_PRICE) + (Val(RSTC!SALES_PRICE) * Val(RSTC!SALES_TAX) / 100)
                    !SALES_TAX = RSTC!SALES_TAX
                    !UNIT = 1 'Val(TXTUNIT.Text)
                    !Remarks = 1 'Val(TXTUNIT.Text)
                    !ORD_QTY = 0
                    !CST = RSTC!CST
                    !ACT_code = Mid(lstmanufact.List(i), 1, 6)
                    !CREATE_DATE = "13/11/2007"
                    !C_USER_ID = ""
                    !MODIFY_DATE = Null
                    !M_USER_ID = ""
                    !CHECK_FLAG = "Y"
                    !SITEM_CODE = ""
                    .Update
                End If
            Next i
            RSTC.Close
            
            Set RSTC = Nothing
        End If
        .Close
        
    End With
    
    RSTB.Close
    
    
    Set RSTA = Nothing
    Set RSTB = Nothing
    
    DataList2_Click
    LSTDISTI.SetFocus
       
    Exit Sub
    
eRRHAND:
    MsgBox Err.Description
    
End Sub

Private Sub cmdadjust_Click()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    If DataList2.BoundText = "" Then Exit Sub
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT [BAL_QTY] from [RTRXFILE] WHERE RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "' AND RTRXFILE.BAL_QTY > 0 ", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTITEMMAST
        If Not (.EOF And .BOF) Then
            !OPEN_QTY = i
            !OPEN_VAL = 0
            !RCPT_QTY = 0
            !RCPT_VAL = 0
            !ISSUE_QTY = 0
            !ISSUE_VAL = 0
            !CLOSE_QTY = i
            !CLOSE_VAL = 0
            !DAM_QTY = 0
            !DAM_VAL = 0
            RSTITEMMAST.Update
        End If
    End With
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    MsgBox "STOCK ADJUSTED", vbOKOnly, "STOCK ADJUST...."
    Exit Sub
    
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdClear_Click()
    
    If grdorder.Rows = 1 Then Exit Sub
    If (MsgBox("This will Delete All Entries...ARE YOU SURE ?", vbYesNo, "DELETE....")) = vbNo Then Exit Sub
    
    On Error GoTo eRRHAND
    
    db2.Execute ("DELETE * FROM TmpOrderlist")
    grdorder.Rows = 1
    lstselDist.Clear
             
    tXTMEDICINE.Text = ""
    tXTMEDICINE.SetFocus
    CmbQty.Text = ""
    
    LBLAVSTOCK.Caption = ""
    LBLMINI.Caption = ""
    LBLMANUFACT.Caption = ""
    LBLUNIT.Caption = 0
        
    Exit Sub
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CmdDelete_Click()
    
    Dim i As Integer
    On Error GoTo eRRHAND
    
    If grdorder.Rows = 1 Then Exit Sub
    If MsgBox("ARE YOU SURE YOU WANT TO DELETE SL. No. " & """" & grdorder.TextMatrix(grdorder.Row, 0) & """", vbYesNo, "DELETE.....") = vbNo Then Exit Sub
    db2.Execute ("delete from [TmpOrderlist] where ITEM_CODE= '" & grdorder.TextMatrix(grdorder.Row, 1) & "' AND ACT_CODE= '" & grdorder.TextMatrix(grdorder.Row, 5) & "'")
    grdorder.Tag = grdorder.Row
    Call fillgridwithSelDist
    Call fillSUPPLIERLIST
    
    tXTMEDICINE.Text = ""
    If grdorder.Rows > Val(grdorder.Tag) Then grdorder.Row = grdorder.Tag
    grdorder.SetFocus
    Exit Sub
eRRHAND:
    If Err.Number = 6148 Then
        tXTMEDICINE.SetFocus
    Else
        MsgBox Err.Description
    End If
    
End Sub

Private Sub CMDDELETESTOCK_Click()
    If grdSTOCKLESS.Rows = 1 Then
        tXTMEDICINE.SetFocus
        Exit Sub
    End If
    If MsgBox("Are You Sure You want to Delete " & "*** " & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 2) & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    db2.Execute ("DELETE from [NONRCVD] WHERE NONRCVD.ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'")
    Call fillSTOCKLESSGRID
    tXTMEDICINE.SetFocus
End Sub

Private Sub CMDDELITEM_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    If DataList2.BoundText = "" Then Exit Sub
    On Error GoTo eRRHAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    If i <> 0 Then
        MsgBox "Cannot Delete " & DataList2.Text & " Since Stock is Available", vbCritical, "DELETING ITEM...."
        Exit Sub
    End If
    
    If MsgBox("Are You Sure You want to Delete " & "*** " & DataList2.Text & " ****", vbYesNo, "DELETING ITEM....") = vbNo Then Exit Sub
    db.Execute ("DELETE from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "'")
    db.Execute ("DELETE from [PRODLINK] where PRODLINK.ITEM_CODE = '" & DataList2.BoundText & "'")
    db.Execute ("DELETE from [ITEMMAST] where ITEMMAST.ITEM_CODE = '" & DataList2.BoundText & "'")
    
    tXTMEDICINE.Tag = tXTMEDICINE.Text
    tXTMEDICINE.Text = ""
    tXTMEDICINE.Text = tXTMEDICINE.Tag
    TXTQTY.Text = ""
    Set LSTDISTI.RowSource = Nothing
    MsgBox "ITEM " & DataList2.Text & "DELETED SUCCESSFULLY", vbInformation, "DELETING ITEM...."
    LBLAVSTOCK.Caption = 0
    LBLMINI.Caption = 0
    Exit Sub
   
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_Click()
    Dim i As Long
    Dim N As Long
    
    N = 0
    For i = 0 To lstselDist.ListCount - 1
        If lstselDist.Selected(i) = True Then
            N = N + 1
            Exit For
        End If
    Next i
        
    If N = 0 Then
        MsgBox "Select atleast one Supplier from the list", vbOKOnly, "Purchase Order!!!"
        Exit Sub
    End If
    Call fillgridwithSelDist
    CMDVALUE.Enabled = True
End Sub

Private Sub cmdexit_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub cmditemcreate_Click()
    frmitemmaster.Show
    MDIMAIN.Enabled = False
End Sub

Private Sub CMDMINQTY_Click()
    
End Sub

Private Sub cmdpurchase_Click()
    Me.Enabled = False
    frmpurchase.Show
End Sub

Private Sub CMDNINQTY_Click()
    If FrmCrimedata.DataList2.BoundText = "" Then Exit Sub
    FrmCrimedata.Enabled = False
    FRMREORDER.Show
End Sub

Private Sub CMDORDER_Click()
    FrmCrimedata.Enabled = False
    frmstockless.Show
End Sub
Private Sub CMDREMOVE_Click()
       
    If DataList2.BoundText = "" Then
        Exit Sub
    End If
    
    If LSTDISTI.Text = "" Then
        Exit Sub
    End If

    If MsgBox("ARE YOU SURE YOU WANT TO REMOVE " & LSTDISTI.Text, vbYesNo, "DELETING....") = vbNo Then Exit Sub
    On Error GoTo eRRHAND
      
    db.Execute ("DELETE *  FROM PRODLINK WHERE ITEM_CODE = '" & DataList2.BoundText & "' AND ACT_CODE = '" & LSTDISTI.BoundText & "'")
    DataList2_Click
    LSTDISTI.SetFocus
       
    Exit Sub
    
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdstockcorrect_Click()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Integer
    
    'Exit Sub
    Screen.MousePointer = vbHourglass
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * from [RTRXFILE]", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rststock.EOF
        rststock!MRP = rststock!SALES_PRICE * Val(rststock!UNIT)
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    Screen.MousePointer = vbNormal
    'Exit Sub
    
    i = 0
    'If txtstockcorrect.Text = "" Then Exit Sub
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTITEMMAST.EOF
        i = 0
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT [BAL_QTY] from [RTRXFILE] WHERE RTRXFILE.ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "'  AND RTRXFILE.BAL_QTY > 0 ", db, adOpenForwardOnly
        Do Until rststock.EOF
            i = i + rststock!BAL_QTY
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        RSTITEMMAST!OPEN_QTY = i
        RSTITEMMAST!OPEN_VAL = 0
        RSTITEMMAST!RCPT_QTY = 0
        RSTITEMMAST!RCPT_VAL = 0
        RSTITEMMAST!ISSUE_QTY = 0
        RSTITEMMAST!ISSUE_VAL = 0
        RSTITEMMAST!CLOSE_QTY = i
        RSTITEMMAST!CLOSE_VAL = 0
        RSTITEMMAST!DAM_QTY = 0
        RSTITEMMAST!DAM_VAL = 0
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    Screen.MousePointer = vbNormal
    Exit Sub
    
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Function CALCULATE_ALL()
    Dim RSTRXFILE As ADODB.Recordset
    Dim RSTVALUE As ADODB.Recordset
    
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    LBLVALUE.Caption = ""
    Set RSTVALUE = New ADODB.Recordset
    RSTVALUE.Open "select ITEM_CODE from [TmpOrderlist]", db2, adOpenForwardOnly
    Do Until RSTVALUE.EOF
        Set RSTRXFILE = New ADODB.Recordset
        RSTRXFILE.Open "Select [UNIT],[PTR] From RTRXFILE  WHERE ITEM_CODE = '" & RSTVALUE!ITEM_CODE & "' ORDER BY VCH_DATE", db, adOpenForwardOnly
        If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
            RSTRXFILE.MoveLast
            ' IIf(IsNull(RSTRXFILE!MRP), "", Format(Val(RSTRXFILE!MRP) * Val(txtunit.Text), ".000"))
            LBLVALUE.Caption = Format(Val(LBLVALUE.Caption) + (Val(RSTRXFILE!PTR) * Val(RSTRXFILE!UNIT)), ".00")
        End If
        RSTRXFILE.Close
        Set RSTRXFILE = Nothing
        
        RSTVALUE.MoveNext
    Loop
    RSTVALUE.Close
    Set RSTVALUE = Nothing
    Screen.MousePointer = vbNormal
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Function

Private Function CALCULATE_SEL()
    Dim RSTRXFILE As ADODB.Recordset
    Dim RSTVALUE As ADODB.Recordset
    
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    LBLVALUE.Caption = ""
    Set RSTVALUE = New ADODB.Recordset
    RSTVALUE.Open "select ITEM_CODE from [SELDIST]", db2, adOpenForwardOnly
    Do Until RSTVALUE.EOF
        Set RSTRXFILE = New ADODB.Recordset
        RSTRXFILE.Open "Select [UNIT],[PTR] From RTRXFILE  WHERE ITEM_CODE = '" & RSTVALUE!ITEM_CODE & "' ORDER BY VCH_DATE", db, adOpenForwardOnly
        If Not (RSTRXFILE.EOF And RSTRXFILE.BOF) Then
            RSTRXFILE.MoveLast
            ' IIf(IsNull(RSTRXFILE!MRP), "", Format(Val(RSTRXFILE!MRP) * Val(txtunit.Text), ".000"))
            LBLVALUE.Caption = Format(Val(LBLVALUE.Caption) + (Val(RSTRXFILE!PTR) * Val(RSTRXFILE!UNIT)), ".000")
        End If
        RSTRXFILE.Close
        Set RSTRXFILE = Nothing
        
        RSTVALUE.MoveNext
    Loop
    RSTVALUE.Close
    Set RSTVALUE = Nothing
    Screen.MousePointer = vbNormal
    Exit Function
    
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Function

Private Sub CMDVALUE_Click()
    If lstselDist.SelCount = 0 Then
        Call CALCULATE_ALL
        CMDVALUE.Enabled = False
        Exit Sub
    End If
    Call CALCULATE_SEL
    CMDVALUE.Enabled = False

End Sub

Private Sub Create_Click()
           
    Dim RSTA As ADODB.Recordset
    Dim RSTB As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
   
    If lstselDist.SelCount = 0 Then
        MsgBox "Select The Distributors to be Ordered", , "ORDER"
        lstselDist.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    db2.Execute ("DELETE * FROM SelDist")
    Set RSTB = New ADODB.Recordset
    RSTB.Open "SELECT * From SelDist", db2, adOpenStatic, adLockOptimistic, adCmdText
    For i = 0 To lstselDist.ListCount - 1
        If lstselDist.Selected(i) = True Then
            Set RSTA = New ADODB.Recordset
            RSTA.Open "SELECT * From TmpOrderlist WHERE ACT_CODE = '" & Trim(Mid(lstselDist.List(i), 1, 6)) & "'", db2, adOpenForwardOnly
            Do Until RSTA.EOF
                
                RSTB.AddNew
                RSTB!Or_Product = RSTA!ITEM_NAME
                RSTB!OR_QTY = RSTA!OR_QTY ''''& "x" & RSTA!OR_UNIT
                RSTB!Or_Distrib = RSTA!ACT_NAME
                RSTB!ITEM_CODE = RSTA!ITEM_CODE
                RSTB!Dist_Code = RSTA!ACT_code
                
                RSTA.MoveNext
                RSTB.Update
            Loop
            RSTA.Close
            Set RSTA = Nothing
        End If
    Next i
    
    RSTB.Close
    Set RSTB = Nothing
    
''''    Call cmdReportGenerate_Click
''''
''''    Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
''''
''''    Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
''''    Print #1, "EXIT"
''''    Close #1
''''
''''    '//HERE write the proper path where your command.com file exist
''''    'Shell "C:\WINDOW\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
''''    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide

    'rptPurchase.ReportFileName = App.Path & "\RPTPurchase.RPT"
    'LBLSHOP.Caption = "(" & LBLSHOP.Caption & ")"
    'rptPurchase.Formulas(0) = "Company = '" & LBLSHOP.Caption & "'"
    
    'rptPurchase.Action = 1
    ReportNameVar = App.Path & "\RPTPurchase.RPT"
    Call cmdReportGenerate_Click
    Call GENERATEREPORT
    Screen.MousePointer = vbNormal
    Exit Sub
eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
     
End Sub

Private Sub DataList2_Click()
    
    Dim RSTAVL As ADODB.Recordset
    Dim RSTMAN As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    Dim i As Integer
    Dim N As Integer
    
    On Error GoTo eRRHAND
    
    LBLAVSTOCK.Caption = ""
    LBLMANUFACT.Caption = ""
            
    LSTDUMMY.Clear
    lstmanufact.Clear
    
    'i = 0
    'Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT [BAL_QTY],[UNIT] from [RTRXFILE] WHERE RTRXFILE.ITEM_CODE = '" & DataList2.BoundText & "' AND RTRXFILE.BAL_QTY <> 0 ", db, adOpenForwardOnly
    'Do Until rststock.EOF
    '    i = i + IIf(IsNull(rststock!BAL_QTY), 0, Val(rststock!BAL_QTY))
    '    rststock.MoveNext
    'Loop
    'rststock.Close
    'Set rststock = Nothing
    'LBLAVSTOCK.Caption = i
    
    Set RSTAVL = New ADODB.Recordset
    RSTAVL.Open "SELECT MANUFACTURER, ITEM_NAME, ITEM_CODE, REORDER_QTY, CLOSE_QTY FROM ITEMMAST WHERE ITEM_CODE = '" & DataList2.BoundText & "'", db, adOpenForwardOnly
    With RSTAVL
        If Not (.EOF And .BOF) Then
            LBLMANUFACT.Caption = RSTAVL!MANUFACTURER
            LBLMINI.Caption = RSTAVL!REORDER_QTY
            LBLAVSTOCK.Caption = RSTAVL!CLOSE_QTY
            If TMPFLAG = True Then
                rstTMP.Open "SELECT PRODLINK.UNIT, PRODLINK.ITEM_CODE, PRODLINK.ITEM_NAME, ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST RIGHT JOIN PRODLINK ON ACTMAST.ACT_CODE = PRODLINK.ACT_CODE WHERE ITEM_CODE = '" & RSTAVL!ITEM_CODE & "' ORDER BY ACTMAST.ACT_NAME", db, adOpenForwardOnly
                TMPFLAG = False
            Else
                rstTMP.Close
                rstTMP.Open "SELECT PRODLINK.UNIT, PRODLINK.ITEM_CODE, PRODLINK.ITEM_NAME, ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST RIGHT JOIN PRODLINK ON ACTMAST.ACT_CODE = PRODLINK.ACT_CODE WHERE ITEM_CODE = '" & RSTAVL!ITEM_CODE & "' ORDER BY ACTMAST.ACT_NAME", db, adOpenForwardOnly
                TMPFLAG = False
            End If
            
           
            Set Me.LSTDISTI.RowSource = rstTMP
            LSTDISTI.ListField = "ACT_NAME"
            LSTDISTI.BoundColumn = "ACT_CODE"
            If Not (rstTMP.EOF Or rstTMP.BOF) Then
                LBLUNIT.Caption = rstTMP!UNIT
            End If
            
            i = 0
            Do Until rstTMP.EOF
            
                LSTDUMMY.AddItem (i)
                LSTDUMMY.List(i) = rstTMP!ACT_code
                i = i + 1
                rstTMP.MoveNext
            
            Loop
              
    End If
        .Close
'        REPFLAG = True
    End With
    Set RSTAVL = Nothing
    
    i = 0
    
    Set RSTMAN = New ADODB.Recordset
    RSTMAN.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenForwardOnly
    With RSTMAN
        Do Until .EOF
            
            For N = 0 To LSTDUMMY.ListCount
                If Trim(LSTDUMMY.List(N)) = Trim(!ACT_code) Then GoTo SKIP
            Next N
            lstmanufact.AddItem (i)
            lstmanufact.List(i) = !ACT_code & " " & Trim(!ACT_NAME)
            i = i + 1
SKIP:
        .MoveNext
        Loop
        .Close
    End With
    Set RSTMAN = Nothing

    If Val(LBLAVSTOCK) < Val(LBLMINI) Then
        LBLAV.ForeColor = vbRed
        LBLAVSTOCK.ForeColor = vbRed
    Else
        LBLAV.ForeColor = vbBlue
        LBLAVSTOCK.ForeColor = vbBlue
    End If
    
    Exit Sub
    
eRRHAND:
    If Err.Number = 3021 Then
        LBLUNIT.Caption = 0
        Resume Next
    Else
        MsgBox Err.Description
    End If
    
End Sub


Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.BoundText = "" Then Exit Sub
            LSTDISTI.SetFocus
                        
    End Select
End Sub

Private Sub Form_Activate()
    Call fillORDERGRID
    Call fillSTOCKLESSGRID
    Call fillSUPPLIERLIST
    SEARCHFLAG = False
End Sub

Private Sub Form_Load()
    
    Dim RSTCOMPANY As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT COMP_NAME, HO_NAME FROM COMPINFO", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        LBLSHOP.Caption = RSTCOMPANY!COMP_NAME '& ", " & RSTCOMPANY!HO_NAME
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    PHY.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenForwardOnly
    
    Set Me.CMBSUPPLIER.RowSource = PHY
    CMBSUPPLIER.ListField = "ACT_NAME"
    CMBSUPPLIER.BoundColumn = "ACT_CODE"
    
    CmbQty.Text = ""
    TMPFLAG = True
    REPFLAG = True
    RCVDFLAG = False
    LBLAVSTOCK.Caption = ""
    LBLMINI.Caption = ""
    LBLMANUFACT.Caption = ""
    LBLUNIT.Caption = 0
    CLOSEALL = 1
    
    Me.Width = 15200
    Me.Height = 10000
    M_EDIT = False
    k = 1
    Me.Left = 0
    Me.Top = 0
    Exit Sub
    
eRRHAND:
    MsgBox (Err.Description)
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        PHY.Close
        If TMPFLAG = False Then rstTMP.Close
        If REPFLAG = False Then RSTREP.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
   Cancel = CLOSEALL
End Sub

Private Sub grdorder_Click()
    TXTsample.Visible = False
    CMBSUPPLIER.Visible = False
    grdorder.SetFocus
End Sub

Private Sub grdorder_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Integer
    If grdorder.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdorder.Col
                Case 3
                    TXTsample.Visible = True
                    TXTsample.Top = grdorder.CellTop + 1740
                    TXTsample.Left = grdorder.CellLeft + 50
                    TXTsample.Width = grdorder.CellWidth
                    TXTsample.Text = grdorder.TextMatrix(grdorder.Row, 6)
                    TXTsample.SetFocus
                Case 7
                    TXTsample.Visible = True
                    TXTsample.Top = grdorder.CellTop + 1740
                    TXTsample.Left = grdorder.CellLeft + 50
                    TXTsample.Width = grdorder.CellWidth
                    TXTsample.Text = grdorder.TextMatrix(grdorder.Row, 7)
                    TXTsample.SetFocus
                Case 4
                    CMBSUPPLIER.Visible = True
                    CMBSUPPLIER.Top = grdorder.CellTop + 1740
                    CMBSUPPLIER.Left = grdorder.CellLeft + 50
                   ' CMBSUPPLIER.Width = grdorder.CellWidth
                    CMBSUPPLIER.Text = grdorder.TextMatrix(grdorder.Row, grdorder.Col)
                    CMBSUPPLIER.SetFocus
            End Select
        Case 114
            sitem = UCase(InputBox("Item Name...?", "Purchase Order!!!!"))
            For i = 1 To grdorder.Rows - 1
                    If Mid(grdorder.TextMatrix(i, 1), 1, Len(sitem)) = sitem Then
                        grdorder.Row = i
                        grdorder.TopRow = i
                    Exit For
                End If
            Next i
            grdorder.SetFocus
    End Select
End Sub

Private Sub grdorder_Scroll()
    TXTsample.Visible = False
    CMBSUPPLIER.Visible = False
    grdorder.SetFocus
End Sub

Private Sub grdSTOCKLESS_Click()

    If grdSTOCKLESS.Rows = 1 Then Exit Sub
    Screen.MousePointer = vbHourglass
    tXTMEDICINE.Text = grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 2)
    DataList2.BoundText = grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1)
    Call DataList2_Click
    LSTDISTI.SetFocus
    Screen.MousePointer = vbNormal
    RCVDFLAG = True
    
End Sub

Private Sub LSTDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If LSTDISTI.BoundText = "" Then Exit Sub
            TXTQTY.SetFocus
                        
    End Select
End Sub

Private Sub lstmanufact_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
        Case vbKeyReturn
            CMDADDDIST.SetFocus
                        
    End Select
End Sub

Private Sub lstselDist_Click()
    Call fillgridwithSelDist
    CMDVALUE.Enabled = True
End Sub

Private Sub tXTMEDICINE_Change()
    
    On Error GoTo eRRHAND
   
   'If Len(tXTMEDICINE.Text) < 2 Then Exit Sub
    If SEARCHFLAG = True Then
        If REPFLAG = True Then
            RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenForwardOnly
            REPFLAG = False
        Else
            RSTREP.Close
            RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%'ORDER BY [ITEM_NAME]", db, adOpenForwardOnly
            REPFLAG = False
        End If
        Set Me.DataList2.RowSource = RSTREP
        DataList2.ListField = "ITEM_NAME"
        DataList2.BoundColumn = "ITEM_CODE"
    Else
        If REPFLAG = True Then
            RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenForwardOnly
            REPFLAG = False
        Else
            RSTREP.Close
            RSTREP.Open "Select DISTINCT [ITEM_CODE],[ITEM_NAME] From ITEMMAST  WHERE ITEM_NAME Like '" & Me.tXTMEDICINE.Text & "%' ORDER BY [ITEM_NAME]", db, adOpenForwardOnly
            REPFLAG = False
        End If
        Set Me.DataList2.RowSource = RSTREP
        DataList2.ListField = "ITEM_NAME"
        DataList2.BoundColumn = "ITEM_CODE"
    End If

    Exit Sub
'RSTREP.Close
'TMPFLAG = False
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.VisibleCount = 0 Then Exit Sub
            DataList2.SetFocus
        Case 116
            SEARCHFLAG = True
    End Select

End Sub

Private Sub tXTMEDICINE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTQTY_GotFocus()
    TXTQTY.SelStart = 0
    TXTQTY.SelLength = Len(TXTQTY.Text)
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.Text) = 0 Then
                If DataList2.SelectedItem Then
                   Exit Sub
                Else
                    tXTMEDICINE.SetFocus
                    Exit Sub
               End If
            End If
            CmbQty.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Public Function AlignLeft(vStr As String, vSpace As Integer) As String
    If Len(Trim(vStr)) > vSpace Then '//if the string length is greater than the space you mention
        AlignLeft = Left(vStr, vSpace)  '&"..."
        Exit Function
    End If
    
    AlignLeft = vStr & Space(vSpace - Len(Trim(vStr)))
End Function

Public Function AlignRight(vNumber As String, vSpace As Integer) As String
    AlignRight = Space(vSpace - Len(Trim(vNumber))) & vNumber
End Function

Public Function RepeatString(vStr As String, vSpace As Integer) As String

    Dim X As Integer
    
    For X = 1 To vSpace
        RepeatString = RepeatString & vStr
    Next X
End Function

Private Function fillSUPPLIERLIST()
    Dim RSTDISTI As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    lstselDist.Clear
    Set RSTDISTI = New ADODB.Recordset
    RSTDISTI.Open "select * from [TmpOrderlist] ORDER BY ACT_NAME, ITEM_NAME", db2, adOpenForwardOnly
    Do Until RSTDISTI.EOF
        For i = 0 To lstselDist.ListCount
            If Trim(RSTDISTI!ACT_code) = Trim(Mid(lstselDist.List(i), 1, 6)) Then GoTo SKIP
        Next i
        lstselDist.AddItem Trim(RSTDISTI!ACT_code) & " " & Trim(RSTDISTI!ACT_NAME)
SKIP:
        RSTDISTI.MoveNext
    Loop
    RSTDISTI.Close
    Set RSTDISTI = Nothing
    
    Screen.MousePointer = vbNormal
 Exit Function
    
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox (Err.Description)
        
End Function

Private Function fillgridwithSelDist()

    Dim RSTORDER As ADODB.Recordset
    Dim i As Long
    Dim N As Long
    
    On Error GoTo eRRHAND
    grdorder.Rows = 1
    i = 0
    N = 0
    
    For N = 0 To lstselDist.ListCount - 1
        Set RSTORDER = New ADODB.Recordset
        RSTORDER.Open "SELECT * From TmpOrderlist WHERE ACT_CODE = '" & Trim(Mid(lstselDist.List(N), 1, 6)) & "'", db2, adOpenForwardOnly
        Do Until RSTORDER.EOF
            If lstselDist.Selected(N) = True Then
                i = i + 1
                grdorder.Rows = grdorder.Rows + 1
                grdorder.FixedRows = 1
                grdorder.TextMatrix(i, 0) = i
                grdorder.TextMatrix(i, 1) = RSTORDER!ITEM_CODE
                grdorder.TextMatrix(i, 2) = RSTORDER!ITEM_NAME
                grdorder.TextMatrix(i, 3) = IIf(RSTORDER!OR_UNIT > 1, RSTORDER!OR_QTY & "x" & RSTORDER!OR_UNIT, RSTORDER!OR_QTY)
                grdorder.TextMatrix(i, 4) = RSTORDER!ACT_NAME
                grdorder.TextMatrix(i, 5) = RSTORDER!ACT_code
                grdorder.TextMatrix(i, 6) = RSTORDER!OR_QTY
                grdorder.TextMatrix(i, 7) = RSTORDER!OR_UNIT
            End If
            RSTORDER.MoveNext
        Loop
        RSTORDER.Close
        Set RSTORDER = Nothing
    Next N
 Exit Function
    
eRRHAND:
    MsgBox (Err.Description)
End Function

Private Function fillSTOCKLESSGRID()
    Dim RSTSTKLESS As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    grdSTOCKLESS.Rows = 1
    grdSTOCKLESS.TextMatrix(0, 0) = "SL"
    grdSTOCKLESS.TextMatrix(0, 1) = "ITEM CODE"
    grdSTOCKLESS.TextMatrix(0, 2) = "ITEM NAME"
    grdSTOCKLESS.TextMatrix(0, 3) = "REMARKS"
    
    grdSTOCKLESS.ColWidth(0) = 500
    grdSTOCKLESS.ColWidth(1) = 0
    grdSTOCKLESS.ColWidth(2) = 2800
    grdSTOCKLESS.ColWidth(3) = 1000
        
    grdSTOCKLESS.ColAlignment(0) = 1
    grdSTOCKLESS.ColAlignment(3) = 3
    i = 0
    Set RSTSTKLESS = New ADODB.Recordset
    RSTSTKLESS.Open "SELECT * FROM NONRCVD ORDER BY ITEM_NAME", db2, adOpenForwardOnly
    Do Until RSTSTKLESS.EOF
        i = i + 1
        grdSTOCKLESS.Rows = grdSTOCKLESS.Rows + 1
        grdSTOCKLESS.FixedRows = 1
        grdSTOCKLESS.TextMatrix(i, 0) = i
        grdSTOCKLESS.TextMatrix(i, 1) = RSTSTKLESS!ITEM_CODE
        grdSTOCKLESS.TextMatrix(i, 2) = RSTSTKLESS!ITEM_NAME
        If RSTSTKLESS!Remarks = "IMP" Then
            grdSTOCKLESS.TextMatrix(i, 3) = "*"
        Else
            grdSTOCKLESS.TextMatrix(i, 3) = IIf(IsNull(RSTSTKLESS!Remarks), "", RSTSTKLESS!Remarks)
        End If
        RSTSTKLESS.MoveNext
    Loop
    RSTSTKLESS.Close
    Set RSTSTKLESS = Nothing
 Exit Function
    
eRRHAND:
    MsgBox (Err.Description)
        
End Function

Private Sub cmdReportGenerate_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim rstSUBfile As ADODB.Recordset
    Dim SN As Integer
    
    SN = 0
    On Error GoTo eRRHAND
    '//NOTE : Report file name should never contain blank space.
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)

    Print #1, Chr(13)
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT DISTINCT Or_Distrib From SelDist ORDER BY Or_Distrib", db2, adOpenForwardOnly
    Do Until RSTTRXFILE.EOF
        
        Set RSTCOMPANY = New ADODB.Recordset
        RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001'", db, adOpenForwardOnly
        If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
            Print #1, Chr(27) & Chr(71) & Chr(10) & _
              Space(12) & Chr(14) & Chr(15) & RSTCOMPANY!COMP_NAME & Space(10) & _
              Chr(15) & Chr(20) & RSTTRXFILE!Or_Distrib & _
              Chr(27) & Chr(72)
            Print #1, Space(12) & "DL NO. " & RSTCOMPANY!CST
            Print #1, Space(19) & RSTCOMPANY!CST
        End If
        RSTCOMPANY.Close
        Set RSTCOMPANY = Nothing
              
        'Print #1, Chr(27) & Chr(67) & Chr(0) & Space(14) & RepeatString("-", 46)
        Print #1,
        'Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SN", 3) & Space(2) & _
            AlignLeft("INV DATE", 8) & Space(10) & _
            AlignLeft("INV AMT", 7) & _
            Chr(27) & Chr(72)  '//Bold Ends
        Print #1, Chr(27) & Chr(71) & Space(12) & Chr(14) & Chr(15) & AlignLeft("SL", 3) & _
                AlignLeft("ITEM NAME", 20) & Space(6) & _
                AlignLeft("QTY", 6) & _
                Chr(27) & Chr(72)  '//Bold Ends
        Print #1, Space(12) & RepeatString("-", 65)
        SN = 0
        Set rstSUBfile = New ADODB.Recordset
        rstSUBfile.Open "SELECT * From SelDist WHERE Or_Distrib = '" & RSTTRXFILE!Or_Distrib & "' ORDER BY Or_Product", db2, adOpenForwardOnly
        Do Until rstSUBfile.EOF
            SN = SN + 1
            Print #1, Chr(27) & Chr(71) & Space(7) & Chr(14) & Chr(15) & AlignRight(Str(SN), 3) & Space(1) & _
                AlignLeft(rstSUBfile!Or_Product, 23) & _
                AlignLeft("->", 2) & Space(1) & _
                AlignLeft(rstSUBfile!OR_QTY, 7) & _
                Chr(27) & Chr(72) & Chr(15) & Chr(20) '//Bold Ends
            Print #1, Chr(13)
            rstSUBfile.MoveNext
        Loop
        Print #1, Chr(13)
        Print #1, Chr(13)
        Print #1, Chr(13)
        Print #1, Chr(13)

        
        rstSUBfile.Close
        Set rstSUBfile = Nothing
        RSTTRXFILE.MoveNext
   Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    Print #1, Chr(13)
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    'MsgBox "Report file generated at " & App.Path & "\Report.txt" & vbCrLf & "Click Print Report Button to print on paper."
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub


Private Function fillORDERGRID()

    Dim RSTORDER As ADODB.Recordset
    Dim i As Integer
    
    'On Error GoTo eRRHAND
    grdorder.Rows = 1
    grdorder.TextMatrix(0, 0) = "SL"
    grdorder.TextMatrix(0, 1) = "ITEM CODE"
    grdorder.TextMatrix(0, 2) = "ITEM NAME"
    grdorder.TextMatrix(0, 3) = "Qty"
    grdorder.TextMatrix(0, 4) = "SUPPLIER"
    grdorder.TextMatrix(0, 5) = "ACT_CODE"
    grdorder.TextMatrix(0, 6) = "OR_QTY"
    grdorder.TextMatrix(0, 7) = "PACK"
    
    grdorder.ColWidth(0) = 500
    grdorder.ColWidth(1) = 0
    grdorder.ColWidth(2) = 2000
    grdorder.ColWidth(3) = 800
    grdorder.ColWidth(4) = 2400
    grdorder.ColWidth(5) = 0
    grdorder.ColWidth(6) = 0
    grdorder.ColWidth(7) = 800
        
    grdorder.ColAlignment(0) = 1
    grdorder.ColAlignment(3) = 3
    grdorder.ColAlignment(7) = 3
    i = 0
    
    Set RSTORDER = New ADODB.Recordset
    RSTORDER.Open "SELECT * FROM TMPORDERLIST ORDER BY ACT_NAME", db2, adOpenForwardOnly
    Do Until RSTORDER.EOF
        i = i + 1
        grdorder.Rows = grdorder.Rows + 1
        grdorder.FixedRows = 1
        grdorder.TextMatrix(i, 0) = i
        grdorder.TextMatrix(i, 1) = RSTORDER!ITEM_CODE
        grdorder.TextMatrix(i, 2) = RSTORDER!ITEM_NAME
        grdorder.TextMatrix(i, 3) = IIf(RSTORDER!OR_UNIT > 1, RSTORDER!OR_QTY & "x" & RSTORDER!OR_UNIT, RSTORDER!OR_QTY)
        grdorder.TextMatrix(i, 4) = RSTORDER!ACT_NAME
        grdorder.TextMatrix(i, 5) = RSTORDER!ACT_code
        grdorder.TextMatrix(i, 6) = RSTORDER!OR_QTY
        grdorder.TextMatrix(i, 7) = RSTORDER!OR_UNIT
        RSTORDER.MoveNext
    Loop
    RSTORDER.Close
    Set RSTORDER = Nothing
 Exit Function
    
eRRHAND:
    MsgBox (Err.Description)
        
End Function

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim RSTORDER As ADODB.Recordset
    Dim M_STOCK As Integer
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdorder.Col
                Case 7   'Pack
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set RSTORDER = New ADODB.Recordset
                    RSTORDER.Open "SELECT OR_UNIT from [TMPORDERLIST] Where ITEM_CODE = '" & grdorder.TextMatrix(grdorder.Row, 1) & "' and ACT_CODE = '" & grdorder.TextMatrix(grdorder.Row, 5) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTORDER.EOF And RSTORDER.BOF) Then
                        RSTORDER!OR_UNIT = Val(TXTsample.Text)
                        RSTORDER.Update
                    End If
                    RSTORDER.Close
                    Set RSTEXP = Nothing
                     
                    grdorder.TextMatrix(grdorder.Row, 7) = Val(TXTsample.Text)
                    grdorder.TextMatrix(grdorder.Row, 3) = grdorder.TextMatrix(grdorder.Row, 6) & "x" & Val(TXTsample.Text)
                    TXTsample.Visible = False
                    grdorder.SetFocus
                Case 3   'QTY
                    If Val(TXTsample.Text) = 0 Then Exit Sub
                    Set RSTORDER = New ADODB.Recordset
                    RSTORDER.Open "SELECT OR_QTY from [TMPORDERLIST] Where ITEM_CODE = '" & grdorder.TextMatrix(grdorder.Row, 1) & "' and ACT_CODE = '" & grdorder.TextMatrix(grdorder.Row, 5) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
                    If Not (RSTORDER.EOF And RSTORDER.BOF) Then
                        RSTORDER!OR_QTY = Val(TXTsample.Text)
                        RSTORDER.Update
                    End If
                    RSTORDER.Close
                    Set RSTEXP = Nothing
                    grdorder.TextMatrix(grdorder.Row, 6) = Val(TXTsample.Text)
                    grdorder.TextMatrix(grdorder.Row, grdorder.Col) = TXTsample.Text & "x" & grdorder.TextMatrix(grdorder.Row, 7)
                    TXTsample.Visible = False
                    grdorder.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdorder.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
    CMBSUPPLIER.Visible = False
End Sub
