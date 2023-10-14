VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMEXPIRY 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EXPIRY"
   ClientHeight    =   9525
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14355
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMEXPIRY.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9525
   ScaleWidth      =   14355
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Height          =   810
      Left            =   6990
      TabIndex        =   26
      Top             =   8655
      Width           =   2850
      Begin VB.CommandButton CMDEXPRET 
         Caption         =   "EXPIRY RETURN"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   75
         TabIndex        =   10
         Top             =   195
         Width           =   1515
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1635
         TabIndex        =   11
         Top             =   195
         Width           =   1110
      End
   End
   Begin MSDataGridLib.DataGrid GrdEXPIRY 
      Height          =   6795
      Left            =   7005
      TabIndex        =   5
      Top             =   225
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   11986
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   14024661
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CMDEXPADD 
      Caption         =   "ADD"
      Height          =   480
      Left            =   10020
      TabIndex        =   6
      Top             =   8865
      Width           =   1005
   End
   Begin VB.CommandButton CMDEDITEXPLIST 
      Caption         =   "ED&IT"
      Height          =   465
      Left            =   11190
      TabIndex        =   7
      Top             =   8880
      Width           =   1020
   End
   Begin MSDataGridLib.DataGrid grdEXPIRYLIST 
      Height          =   7530
      Left            =   60
      TabIndex        =   3
      Top             =   1125
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   13282
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16768407
      ForeColor       =   21760
      HeadLines       =   1
      RowHeight       =   15
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
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame FRMEEXPLIST 
      BackColor       =   &H00C0C0FF&
      Height          =   825
      Left            =   9870
      TabIndex        =   25
      Top             =   8640
      Width           =   4665
      Begin VB.CommandButton CMDEXPPRINT 
         Caption         =   "&PRINT"
         Height          =   465
         Left            =   3570
         TabIndex        =   9
         Top             =   225
         Width           =   1035
      End
      Begin VB.CommandButton CMDEXPDELETE 
         Caption         =   "DE&LETE"
         Height          =   465
         Left            =   2460
         TabIndex        =   8
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame FRMEEXPIRY 
      BackColor       =   &H00C0C0FF&
      Height          =   870
      Left            =   60
      TabIndex        =   24
      Top             =   8625
      Width           =   2775
      Begin VB.CommandButton CMDADD 
         Caption         =   "&ADD TO EXPIRY LIST"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   75
         TabIndex        =   4
         Top             =   195
         Width           =   1680
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Height          =   435
         Left            =   1845
         Picture         =   "FRMEXPIRY.frx":030A
         Top             =   255
         Width           =   825
      End
   End
   Begin VB.Frame FRMEDISPLAY 
      BackColor       =   &H00C0C0FF&
      Height          =   915
      Left            =   30
      TabIndex        =   22
      Top             =   -30
      Width           =   5850
      Begin VB.CommandButton CMDPRINT 
         Caption         =   "&PRINT"
         Height          =   435
         Left            =   4230
         TabIndex        =   2
         Top             =   255
         Width           =   1230
      End
      Begin VB.CommandButton CMDDISPLAY 
         Caption         =   "&DISPLAY"
         Height          =   435
         Left            =   2835
         TabIndex        =   1
         Top             =   255
         Width           =   1230
      End
      Begin VB.TextBox TXTDAYS 
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
         Left            =   2265
         MaxLength       =   2
         TabIndex        =   0
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ENTER THE MONTH(s)"
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
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   330
         Width           =   2025
      End
   End
   Begin VB.Frame FRMEPRINT 
      BackColor       =   &H00C0C0FF&
      Height          =   660
      Left            =   7020
      TabIndex        =   37
      Top             =   7980
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox chksent 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Marked as Sent"
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
         Height          =   375
         Left            =   4545
         TabIndex        =   21
         Top             =   180
         Width           =   1800
      End
      Begin MSDataListLib.DataCombo CMBexpdist 
         Height          =   330
         Left            =   855
         TabIndex        =   20
         Top             =   210
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   582
         _Version        =   393216
         Style           =   2
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DISTI"
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
         Left            =   150
         TabIndex        =   38
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame FRMFIELDS 
      BackColor       =   &H00C0C0FF&
      Height          =   1440
      Left            =   7005
      TabIndex        =   27
      Top             =   7185
      Width           =   7515
      Begin VB.TextBox TXTUNIT 
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
         Left            =   6705
         MaxLength       =   3
         TabIndex        =   14
         Top             =   195
         Width           =   615
      End
      Begin VB.TextBox TXTEXPQTY 
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
         Left            =   5325
         MaxLength       =   3
         TabIndex        =   13
         Top             =   210
         Width           =   675
      End
      Begin VB.TextBox TXTBATCH 
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
         Height          =   390
         Left            =   5325
         MaxLength       =   12
         TabIndex        =   16
         Top             =   585
         Width           =   1980
      End
      Begin VB.TextBox TXTEXPDATE 
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
         Left            =   825
         MaxLength       =   10
         TabIndex        =   17
         Top             =   1005
         Width           =   1410
      End
      Begin VB.TextBox TXTITEM 
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
         Left            =   870
         MaxLength       =   50
         TabIndex        =   12
         Top             =   195
         Width           =   3360
      End
      Begin VB.TextBox TXTMFGR 
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
         Left            =   2820
         MaxLength       =   30
         TabIndex        =   18
         Top             =   1020
         Width           =   1455
      End
      Begin VB.TextBox TXTMRP 
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
         Left            =   5010
         MaxLength       =   6
         TabIndex        =   19
         Top             =   1005
         Width           =   690
      End
      Begin MSDataListLib.DataCombo CMBDISTI 
         Height          =   330
         Left            =   870
         TabIndex        =   15
         Top             =   600
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "UNIT"
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
         Index           =   9
         Left            =   6165
         TabIndex        =   39
         Top             =   255
         Width           =   555
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM"
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
         Left            =   195
         TabIndex        =   34
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "QTY"
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
         Left            =   4740
         TabIndex        =   33
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "DISTI"
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
         Left            =   165
         TabIndex        =   32
         Top             =   615
         Width           =   720
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "BATCH"
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
         Left            =   4485
         TabIndex        =   31
         Top             =   615
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "EXPIRY"
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
         Left            =   60
         TabIndex        =   30
         Top             =   1005
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MFGR"
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
         Index           =   6
         Left            =   2265
         TabIndex        =   29
         Top             =   1035
         Width           =   540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "MRP"
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
         Index           =   7
         Left            =   4425
         TabIndex        =   28
         Top             =   1065
         Width           =   555
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXPIRED ITEMS TO BE RETURNED"
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
      Left            =   6990
      TabIndex        =   36
      Top             =   15
      Width           =   7470
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "EXPIRED ITEMS"
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
      Left            =   45
      TabIndex        =   35
      Top             =   870
      Width           =   5895
   End
End
Attribute VB_Name = "FRMEXPIRY"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PHY_FLAG As Boolean
Dim ACT_FLAG As Boolean
Dim CMB_FLAG As Boolean
Dim EXP_FLAG As Boolean
Dim CLOSEALL As Integer

Dim PHY_REC As New ADODB.Recordset
Dim ACT_REC As New ADODB.Recordset
Dim EXP_REC As New ADODB.Recordset
Dim CMB_REC As New ADODB.Recordset

Private Sub chksent_Click()
    Dim RSTFLAG As ADODB.Recordset
    
    If Trim(CMBexpdist.Text) = "" Then
        chksent.Value = 0
        Exit Sub
    End If
    
    On Error GoTo eRRHAND
    
    If chksent.Value = 1 Then
        If (MsgBox("ARE YOU SURE YOU WANT TO MARK " & Trim(CMBexpdist.Text) & " AS SENT", vbYesNo) = vbNo) Then
            chksent.Value = 0
            Exit Sub
        End If
        Set RSTFLAG = New ADODB.Recordset
        RSTFLAG.Open "SELECT * FROM EXPLIST where EXPLIST.EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db2, adOpenStatic, adLockOptimistic, adCmdText
        
        Do Until RSTFLAG.EOF
            RSTFLAG!EX_FLAG = "Y"
            RSTFLAG!EX_SENTDATE = Date
            RSTFLAG.Update
            RSTFLAG.MoveNext
        Loop
        RSTFLAG.Close
        Set RSTFLAG = Nothing
        
        Call FILLCOMBO
        CMBexpdist.Text = ""
        Call FILLGRID
        chksent.Value = 0
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description

End Sub

Private Sub CMBDISTI_GotFocus()
    CMBDISTI.SelStart = 0
    CMBDISTI.SelLength = Len(CMBDISTI.Text)
End Sub

Private Sub CMBDISTI_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(CMBDISTI.Text) = "" Then Exit Sub
            txtBatch.SetFocus
                        
    End Select
End Sub

Private Sub CMBDISTI_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub CMBexpdist_Change()
    Call FILLGRID
End Sub

Private Sub CMBexpdist_Click(Area As Integer)
    '
End Sub

Private Sub CMDADD_Click()
    Dim RSTEXP As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    
    Dim i As Integer
    Dim M_DATA As Integer
    Dim K_DATA As Integer
    Dim FCODE As Integer
    
    Dim FCODE1, FCODE2, FCODE3, FCODE4, FCODE5, FCODE6, FCODE7, FCODE8, FCODE9, FCODE10, FCODE11, FCODE12, FCODE13, FCODE14   As String

    On Error GoTo eRRHAND
    
    If grdEXPIRYLIST.ApproxCount < 1 Then Exit Sub
    
    FCODE = GrdEXPIRY.ApproxCount + 1
    FCODE1 = grdEXPIRYLIST.Columns(1) 'ITEM
    FCODE2 = grdEXPIRYLIST.Columns(2) 'INVOICE
    FCODE3 = grdEXPIRYLIST.Columns(3) 'PURCHASE DATE
    FCODE4 = grdEXPIRYLIST.Columns(4) ' MFGR
    FCODE5 = grdEXPIRYLIST.Columns(5) ' DISTI
    FCODE6 = grdEXPIRYLIST.Columns(6) 'BATCH
    FCODE7 = grdEXPIRYLIST.Columns(7)
    FCODE8 = grdEXPIRYLIST.Columns(8) 'QTY
    FCODE9 = grdEXPIRYLIST.Columns(9) ' MRP
    FCODE10 = grdEXPIRYLIST.Columns(10) 'SETTLE
    FCODE11 = grdEXPIRYLIST.Columns(11) 'VALUE
    FCODE12 = "N"   'FLAG
    FCODE13 = grdEXPIRYLIST.Columns(15) 'UNIT
    FCODE14 = Date  'SEND DATE
    
    
    'db2.Execute ("Insert into EXPLIST values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "','" & FCODE7 & "','" & FCODE8 & "','" & FCODE9 & "','" & FCODE10 & "','" & FCODE11 & "','" & FCODE12 & "','" & FCODE13 & "','" & FCODE14 & "' )")
    db2.Execute ("Delete from [EXPIRY] where EXPIRY.EX_SLNO = '" & Val(grdEXPIRYLIST.Columns(0)) & "'")
    
    M_DATA = 0
    K_DATA = 0
    Set RSTEXP = New ADODB.Recordset
    RSTEXP.Open "SELECT * from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & grdEXPIRYLIST.Columns(14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    
    Do Until RSTEXP.EOF
        If Not (RSTEXP!LINE_NO = Val(grdEXPIRYLIST.Columns(13)) And RSTEXP!VCH_NO = Val(grdEXPIRYLIST.Columns(12))) Then GoTo SKIP
        Set RSTITEM = New ADODB.Recordset
        RSTITEM.Open "SELECT * from [ITEMMAST] where ITEMMAST.ITEM_CODE = '" & grdEXPIRYLIST.Columns(14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        
        M_DATA = Val(RSTITEM!ISSUE_QTY) + Val(grdEXPIRYLIST.Columns(8).Text)
        RSTITEM!ISSUE_QTY = M_DATA
        
        RSTITEM!ISSUE_VAL = 0
        
        K_DATA = RSTITEM!CLOSE_QTY - Val(grdEXPIRYLIST.Columns(8).Text)
        RSTITEM!CLOSE_QTY = K_DATA
        
        RSTITEM!CLOSE_VAL = 0
        RSTEXP!ISSUE_QTY = RSTEXP!BAL_QTY
        RSTEXP!BAL_QTY = 0
        RSTITEM.Update
        RSTEXP.Update
        RSTITEM.Close
        Set RSTITEM = Nothing
SKIP:
        RSTEXP.MoveNext
    Loop
    
    RSTEXP.Close
    
    Set RSTEXP = Nothing
    
    Call FILLEXPIRYGRID
    Call FILLEXPIRYLIST
    Screen.MousePointer = vbNormal
    
    Exit Sub
   
eRRHAND:
    If Err.Number = 7005 Then
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CMDDISPLAY_Click()
    
    Dim RSTD As ADODB.Recordset
    Dim RSTE As ADODB.Recordset
    Dim RSTF As ADODB.Recordset
    
    Dim M_DATE As Date
    Dim E_DATE As Date
    
    Dim i As Integer
    
    
    If FRMEXPIRY.TXTDAYS.Text = "" Then Exit Sub
    
    On Error GoTo eRRHAND
    
    i = LastDayOfMonth(Date)
    M_DATE = i & "/" & Month(Date) & "/" & Year(Date)
    
    E_DATE = DateAdd("m", Val(TXTDAYS.Text), M_DATE)
    
    

    Screen.MousePointer = vbHourglass
    db2.Execute ("DELETE * FROM EXPIRY")
    
    i = 0
    Set RSTD = New ADODB.Recordset
    RSTD.Open "SELECT * From EXPIRY", db2, adOpenStatic, adLockOptimistic, adCmdText
    Set RSTE = New ADODB.Recordset
    RSTE.Open "SELECT * From [RTRXFILE] WHERE [BAL_QTY] <> 0 AND [EXP_DATE] <=# " & E_DATE & " # ORDER BY RTRXFILE.EXP_DATE", db, adOpenStatic, adLockReadOnly, adCmdText
   
    Do Until RSTE.EOF
        
        RSTD.AddNew
        i = i + 1
        RSTD!EX_SLNO = i
        RSTD!EX_ITEM = RSTE!ITEM_NAME
        RSTD!EX_PUR_INV = RSTE!PINV
        RSTD!EX_PUR_DATE = RSTE!CREATE_DATE
        
        Set RSTF = New ADODB.Recordset
        RSTF.Open "SELECT * From ITEMMAST WHERE ITEM_CODE ='" & RSTE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
        
        RSTD!EX_MFGR = RSTF!MANUFACTURER
        If Len(RSTE!VCH_DESC) > 14 Then
            If Mid(RSTE!VCH_DESC, 1, 14) = "Received From " Then
                RSTD!EX_DISTI = Mid(RSTE!VCH_DESC, 15)
            Else
                RSTD!EX_DISTI = RSTE!VCH_DESC
            End If
        Else
            RSTD!EX_DISTI = RSTE!VCH_DESC
        End If
        RSTD!EX_BATCH = RSTE!REF_NO
        RSTD!EX_DATE = Format(RSTE!EXP_DATE, "MM/YY")
        RSTD!EX_QTY = RSTE!BAL_QTY
        RSTD!EX_MRP = RSTE!MRP
        RSTD!EX_SETTLE = ""
        RSTD!EX_VALUE = Val(RSTE!MRP) * Val(RSTE!BAL_QTY)
        RSTD!VCH_NO = RSTE!VCH_NO
        RSTD!LINE_NO = RSTE!LINE_NO
        RSTD!ITEM_CODE = RSTE!ITEM_CODE
        RSTD!EX_UNIT = RSTE!UNIT
        RSTD!PINV = RSTE!PINV
        
        RSTD.Update
        
        RSTF.Close
        Set RSTF = Nothing
            
        RSTE.MoveNext
            
    Loop

    RSTE.Close
 
     Set RSTE = Nothing

    
    RSTD.Close
    Set RSTD = Nothing
    
    Call FILLEXPIRYLIST
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
     
End Sub

Private Sub cmdedit_Click()

    Dim RSTEXP As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    
    Dim i As Integer
    Dim M_DATA As Integer
    Dim K_DATA As Integer
    Dim FCODE As Integer
    
    Dim FCODE1, FCODE2, FCODE3, FCODE4, FCODE5, FCODE6, FCODE7, FCODE8, FCODE9, FCODE10, FCODE11, FCODE12, FCODE13, FCODE14   As String

    On Error GoTo eRRHAND
    If grdEXPIRYLIST.ApproxCount < 1 Then Exit Sub
    
    FCODE = GrdEXPIRY.ApproxCount + 1
    FCODE1 = grdEXPIRYLIST.Columns(1)
    FCODE2 = grdEXPIRYLIST.Columns(2)
    FCODE3 = grdEXPIRYLIST.Columns(3)
    FCODE4 = grdEXPIRYLIST.Columns(4)
    FCODE5 = grdEXPIRYLIST.Columns(5)
    FCODE6 = grdEXPIRYLIST.Columns(6)
    FCODE7 = grdEXPIRYLIST.Columns(7)
    FCODE8 = grdEXPIRYLIST.Columns(8)
    FCODE9 = grdEXPIRYLIST.Columns(9)
    FCODE10 = grdEXPIRYLIST.Columns(10)
    FCODE11 = grdEXPIRYLIST.Columns(11)
    FCODE12 = "N"
    'FCODE13 = grdEXPIRYLIST.Columns(13)
    'FCODE14 = DATE
    
    
    
    db2.Execute ("Insert into EXPLIST values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "','" & FCODE7 & "','" & FCODE8 & "','" & FCODE9 & "','" & FCODE10 & "','" & FCODE11 & "' )")
    db2.Execute ("Delete from [EXPIRY] where EXPIRY.EX_SLNO = '" & Val(grdEXPIRYLIST.Columns(0)) & "'")
    
    M_DATA = 0
    K_DATA = 0
    Set RSTEXP = New ADODB.Recordset
    'RSTEXP.Open "SELECT * from [RTRXFILE] where RTRXFILE.ITEM_CODE = 'grdEXPIRYLIST.Columns(14)' AND RTRXFILE.VCH_NO = 'grdEXPIRYLIST.Columns(12)' AND RTRXFILE.LINE_NO = 'grdEXPIRYLIST.Columns(13)'", db, adOpenStatic, adLockReadOnly
    RSTEXP.Open "SELECT * from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & grdEXPIRYLIST.Columns(14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    
    Do Until RSTEXP.EOF
        If Not (RSTEXP!LINE_NO = Val(grdEXPIRYLIST.Columns(13)) And RSTEXP!VCH_NO = Val(grdEXPIRYLIST.Columns(12))) Then GoTo SKIP
        Set RSTITEM = New ADODB.Recordset
        RSTITEM.Open "SELECT * from [ITEMMAST] where ITEMMAST.ITEM_CODE = '" & grdEXPIRYLIST.Columns(14) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        
        M_DATA = Val(RSTITEM!ISSUE_QTY) + Val(grdEXPIRYLIST.Columns(8).Text)
        RSTITEM!ISSUE_QTY = M_DATA
        
        RSTITEM!ISSUE_VAL = 0
        
        K_DATA = RSTITEM!CLOSE_QTY - Val(grdEXPIRYLIST.Columns(8).Text)
        RSTITEM!CLOSE_QTY = K_DATA
        
        RSTITEM!CLOSE_VAL = 0
        RSTEXP!ISSUE_QTY = RSTEXP!BAL_QTY
        RSTEXP!BAL_QTY = 0
        RSTITEM.Update
        RSTEXP.Update
        RSTITEM.Close
        Set RSTITEM = Nothing
SKIP:
        RSTEXP.MoveNext
    Loop
    
    RSTEXP.Close
    
    Set RSTEXP = Nothing
    
    Call FILLEXPIRYGRID
    Call FILLEXPIRYLIST
    Screen.MousePointer = vbNormal
    
    Exit Sub
   
eRRHAND:
    If Err.Number = 7005 Then
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CMDEDITEXPLIST_Click()
    
    Dim RSTEXP As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo eRRHAND
    
    If GrdEXPIRY.ApproxCount < 1 Then Exit Sub

    If CMDEDITEXPLIST.Caption = "ED&IT" Then
        
        'If MsgBox("ARE YOU SURE YOU WANT TO EDIT " & """" & GrdEXPIRY.Columns(1) & """", vbYesNo, "EDIT....") = vbNo Then Exit Sub
        
        GrdEXPIRY.Height = 6960
        FRMFIELDS.Enabled = True
        CMBDISTI.Text = GrdEXPIRY.Columns(5)
        txtBatch.Text = GrdEXPIRY.Columns(6)
        TXTEXPDATE = GrdEXPIRY.Columns(7)
        TXTEXPQTY = GrdEXPIRY.Columns(8)
        TXTUNIT = GrdEXPIRY.Columns(13)
        TXTITEM.Text = GrdEXPIRY.Columns(1)
        TXTITEM.SetFocus
        TXTMFGR.Text = GrdEXPIRY.Columns(4)
        TxtMRP.Text = GrdEXPIRY.Columns(9)
        
        
        FRMEEXPLIST.Enabled = False
        FRMEDISPLAY.Enabled = False
        grdEXPIRYLIST.Enabled = False
        GrdEXPIRY.Enabled = False
        FRMEEXPIRY.Enabled = False
        CMDEDITEXPLIST.Caption = "&SAVE"
        CMDEXIT.Caption = "CANCEL"
        CMDEXPADD.Enabled = False
        CMDEXPPRINT.Enabled = False
    Else
        If Trim(TXTITEM.Text) = "" Then
            MsgBox "ENTER THE ITEM", vbOKOnly, "EXPIRY"
            TXTITEM.SetFocus
            Exit Sub
        End If
        
        If Val(TXTEXPQTY.Text) = 0 Then
            MsgBox "ENTER THE QTY", vbOKOnly, "EXPIRY"
            TXTEXPQTY.SetFocus
            Exit Sub
        End If
        
        If Val(TXTUNIT.Text) = 0 Then
            MsgBox "ENTER THE UNIT", vbOKOnly, "EXPIRY"
            TXTUNIT.SetFocus
            Exit Sub
        End If
        
        If Trim(CMBDISTI.Text) = "" Then
            MsgBox "ENTER THE NAME OF DISTRIBUTOR", vbOKOnly, "EXPIRY"
            CMBDISTI.SetFocus
            Exit Sub
        End If
              
        If Trim(TXTEXPDATE.Text) <> "" Then
            If Not IsDate(TXTEXPDATE.Text) Then
                MsgBox "ENTER A VALID DATE FOR EXPIRY", vbOKOnly, "EXPIRY"
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
        End If
        
        Set RSTEXP = New ADODB.Recordset
        RSTEXP.Open "SELECT * from [EXPLIST] where EXPLIST.EX_SLNO = '" & GrdEXPIRY.Columns(0) & "'", db2, adOpenStatic, adLockOptimistic, adCmdText
    
        If Not (RSTEXP.EOF And RSTEXP.BOF) Then
            RSTEXP!EX_DISTI = Trim(CMBDISTI.Text)
            RSTEXP!EX_BATCH = Trim(txtBatch.Text)
            RSTEXP!EX_DATE = Trim(TXTEXPDATE.Text)
            RSTEXP!EX_QTY = Val(TXTEXPQTY.Text)
            RSTEXP!EX_ITEM = Trim(TXTITEM.Text)
            RSTEXP!EX_MFGR = Trim(TXTMFGR.Text)
            RSTEXP!EX_MRP = Val(TxtMRP.Text)
            RSTEXP!EX_UNIT = Val(TXTUNIT.Text)
            RSTEXP!EX_SENTDATE = Date
            RSTEXP.Update
        End If
        RSTEXP.Close
        
        Set RSTEXP = Nothing
        CMBDISTI.Text = ""
        txtBatch.Text = ""
        TXTEXPDATE = ""
        TXTEXPQTY = ""
        TXTUNIT = ""
        TXTITEM.Text = ""
        TXTMFGR.Text = ""
        TxtMRP.Text = ""
        
        FRMEEXPLIST.Enabled = True
        FRMEDISPLAY.Enabled = True
        grdEXPIRYLIST.Enabled = True
        GrdEXPIRY.Enabled = True
        FRMEEXPIRY.Enabled = True
        CMDEDITEXPLIST.Caption = "ED&IT"
        GrdEXPIRY.Height = 6960
        FRMFIELDS.Enabled = False
        CMDEXPADD.Enabled = True
        CMDEXPPRINT.Enabled = True
        CMDEXIT.Caption = "E&XIT"
        Call FILLEXPIRYGRID
        MsgBox "Saved Successfully", vbOKOnly, "EXPIRY"
    End If
    
    Screen.MousePointer = vbNormal
    
    Exit Sub
   
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    If CMDEXIT.Caption = "CANCEL" Then
        CMDEXIT.Caption = "E&XIT"
        CMBDISTI.Text = ""
        txtBatch.Text = ""
        TXTEXPDATE = ""
        TXTEXPQTY = ""
        TXTUNIT = ""
        TXTITEM.Text = ""
        TXTMFGR.Text = ""
        TxtMRP.Text = ""
        
        FRMEEXPLIST.Enabled = True
        FRMEDISPLAY.Enabled = True
        grdEXPIRYLIST.Enabled = True
        GrdEXPIRY.Enabled = True
        FRMEEXPIRY.Enabled = True
        CMDEDITEXPLIST.Caption = "ED&IT"
        CMDEXPADD.Caption = "ADD"
        CMDEXPPRINT.Caption = "&PRINT"
        GrdEXPIRY.Height = 8400
        FRMFIELDS.Visible = True
        FRMFIELDS.Enabled = False
        CMDEXPADD.Enabled = True
        FRMEPRINT.Visible = False
        CMDEDITEXPLIST.Enabled = True
        CMDEXPPRINT.Enabled = True
        
        Call FILLEXPIRYGRID
        CLOSEALL = 1
        Exit Sub
    End If
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub CMDEXPADD_Click()
    Dim FCODE, FCODE1, FCODE2, FCODE3, FCODE4, FCODE5, FCODE6, FCODE7, FCODE8, FCODE9, FCODE11, FCODE10, FCODE12, FCODE13, FCODE14 As String
    Dim i As Integer
    
    On Error GoTo eRRHAND
        
    If CMDEXPADD.Caption = "ADD" Then
        GrdEXPIRY.Height = 6960
        FRMFIELDS.Enabled = True
        FRMEEXPLIST.Enabled = False
        FRMEDISPLAY.Enabled = False
        grdEXPIRYLIST.Enabled = False
        GrdEXPIRY.Enabled = False
        FRMEEXPIRY.Enabled = False
        CMDEXPADD.Caption = "&SAVE"
        CMDEXIT.Caption = "CANCEL"
        CMDEDITEXPLIST.Enabled = False
        CMDEXPPRINT.Enabled = False
        TXTITEM.SetFocus
    Else
    
        If Trim(TXTITEM.Text) = "" Then
            MsgBox "ENTER THE ITEM", vbOKOnly, "EXPIRY"
            TXTITEM.SetFocus
            Exit Sub
        End If
        
        If Val(TXTEXPQTY.Text) = 0 Then
            MsgBox "ENTER THE QTY", vbOKOnly, "EXPIRY"
            TXTEXPQTY.SetFocus
            Exit Sub
        End If
        
        If Val(TXTUNIT.Text) = 0 Then
            MsgBox "ENTER THE UNIT", vbOKOnly, "EXPIRY"
            TXTUNIT.SetFocus
            Exit Sub
        End If
        
        If Trim(CMBDISTI.Text) = "" Then
            MsgBox "ENTER THE NAME OF DISTRIBUTOR", vbOKOnly, "EXPIRY"
            CMBDISTI.SetFocus
            Exit Sub
        End If
              
        If Trim(TXTEXPDATE.Text) <> "" Then
            If Not IsDate(TXTEXPDATE.Text) Then
                MsgBox "ENTER A VALID DATE FOR EXPIRY", vbOKOnly, "EXPIRY"
                TXTEXPDATE.SetFocus
                Exit Sub
            End If
        End If
        
        FCODE = GrdEXPIRY.ApproxCount + 1
        FCODE1 = Trim(TXTITEM.Text)
        FCODE2 = ""   'INVOICE
        FCODE3 = ""   'INV DATE
        FCODE4 = Trim(TXTMFGR.Text)
        FCODE5 = Trim(CMBDISTI.Text)
        FCODE6 = Trim(txtBatch.Text)
        FCODE7 = Trim(TXTEXPDATE.Text)
        FCODE8 = Val(TXTEXPQTY.Text)
        FCODE9 = Val(TxtMRP.Text)
        FCODE10 = ""   'SETTLE
        FCODE11 = Val(TxtMRP.Text) * Val(TXTEXPQTY.Text) 'VALUE
        FCODE12 = "N"
        FCODE13 = Val(TXTUNIT.Text)
        FCODE14 = Date
        
        db2.Execute ("Insert into EXPLIST values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "','" & FCODE7 & "','" & FCODE8 & "','" & FCODE9 & "','" & FCODE10 & "','" & FCODE11 & "','" & FCODE12 & "','" & FCODE13 & "','" & FCODE14 & "' )")
        CMBDISTI.Text = ""
        txtBatch.Text = ""
        TXTEXPDATE = ""
        TXTEXPQTY = ""
        TXTUNIT = ""
        TXTITEM.Text = ""
        TXTMFGR.Text = ""
        TxtMRP.Text = ""
        
        FRMEEXPLIST.Enabled = True
        FRMEDISPLAY.Enabled = True
        grdEXPIRYLIST.Enabled = True
        GrdEXPIRY.Enabled = True
        FRMEEXPIRY.Enabled = True
        CMDEXPADD.Caption = "ADD"
        GrdEXPIRY.Height = 8400
        FRMFIELDS.Enabled = False
        CMDEXIT.Caption = "E&XIT"
        CMDEDITEXPLIST.Enabled = True
        CMDEXPPRINT.Enabled = True
        Call FILLEXPIRYGRID
        MsgBox "Saved Successfully", vbOKOnly, "EXPIRY"
    End If
    Screen.MousePointer = vbNormal
    
    Exit Sub
   
eRRHAND:
    MsgBox Err.Description
End Sub

Private Sub CMDEXPDELETE_Click()
    
    On Error GoTo eRRHAND
    If GrdEXPIRY.ApproxCount > 0 Then
        If MsgBox("ARE YOU SURE YOU WANT TO DELETE " & """" & GrdEXPIRY.Columns(1) & """", vbYesNo, "EDIT....") = vbNo Then Exit Sub
        Screen.MousePointer = vbHourglass
        db2.Execute ("Delete from [EXPLIST] where EX_SLNO= '" & GrdEXPIRY.Columns(0) & "' AND EX_FLAG ='N'")
        Call FILLEXPIRYGRID
    End If
    Exit Sub

eRRHAND:
     MsgBox Err.Description
End Sub

Private Sub CMDEXPPRINT_Click()
    Dim RSTEXP As ADODB.Recordset
    Dim RSTSORT As ADODB.Recordset
    
    On Error GoTo eRRHAND
    
    If GrdEXPIRY.ApproxCount < 1 Then Exit Sub
    
    If CMDEXPPRINT.Caption = "&PRINT" Then
        GrdEXPIRY.Height = 7755
        FRMFIELDS.Visible = False
        FRMEPRINT.Visible = True
        CMDEXPPRINT.Caption = "OK"
        CMDEXIT.Caption = "CANCEL"
        
        CMDEXPADD.Enabled = False
        CMDEDITEXPLIST.Enabled = False
        
        Call FILLCOMBO
        CMBexpdist.Text = ""
    Else
        
        Call FILLCOMBO
        
        Call cmdReportGenerate_Click
    
        Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
    
        Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
        Print #1, "EXIT"
        Close #1
    
    '//HERE write the proper path where your command.com file exist
    'Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    Shell "C:\WINDOWS\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
   
eRRHAND:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
    
End Sub

Private Sub CMDEXPRET_Click()
    Me.Enabled = False
    FRMEXPRCVD.Show
End Sub

Private Sub cmdprint_Click()
    
    If grdEXPIRYLIST.VisibleRows < 2 Then Exit Sub
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    Call EXPIRYReport
    Open App.Path & "\repo.bat" For Output As #1 '//Creating Batch file
    
    Print #1, "TYPE " & App.Path & "\Report.txt > PRN"
    Print #1, "EXIT"
    Close #1
    
    '//HERE write the proper path where your command.com file exist
    Shell "C:\WINDOWS\COMMAND.COM /C " & App.Path & "\REPO.BAT N", vbHide
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Load()
    Dim RSTACT As ADODB.Recordset
    Dim N As Integer
    
    'FrmCrimedata.Enabled = False
    PHY_FLAG = True
    EXP_FLAG = True
    ACT_FLAG = True
    CMB_FLAG = True
    
    CLOSEALL = 1
    
    On Error GoTo eRRHAND
    
    Call FILLEXPIRYGRID
    
    Set CMBDISTI.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set Me.CMBDISTI.RowSource = ACT_REC
    CMBDISTI.ListField = "ACT_NAME"
    CMBDISTI.BoundColumn = "ACT_CODE"
    GrdEXPIRY.Height = 8400
    FRMFIELDS.Enabled = False
    Me.Left = 0
    Me.Top = 0
    Me.Height = 10000
    Me.Width = 14800
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        If EXP_FLAG = False Then EXP_REC.Close
        If PHY_FLAG = False Then PHY_REC.Close
        If ACT_FLAG = False Then ACT_REC.Close
        If CMB_FLAG = False Then CMB_REC.Close
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub TXTBATCH_GotFocus()
    txtBatch.SelStart = 0
    txtBatch.SelLength = Len(txtBatch.Text)
End Sub

Private Sub TXTBATCH_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(TXTBATCH.Text) = "" Then Exit Sub
            TXTEXPDATE.SetFocus
                        
    End Select
End Sub

Private Sub TXTBATCH_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("-")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTDAYS_GotFocus()
    TXTDAYS.SelStart = 0
    TXTDAYS.SelLength = Len(TXTDAYS.Text)
End Sub

Private Sub TXTDAYS_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTDAYS.Text) = 0 Then Exit Sub
            CMDDISPLAY.SetFocus
                        
    End Select
End Sub

Private Sub TXTDAYS_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Function LastDayOfMonth(DateIn)
    Dim TempDate
    TempDate = Year(DateIn) & "-" & Month(DateIn) & "-"
    If IsDate(TempDate & "28") Then LastDayOfMonth = 28
    If IsDate(TempDate & "29") Then LastDayOfMonth = 29
    If IsDate(TempDate & "30") Then LastDayOfMonth = 30
    If IsDate(TempDate & "31") Then LastDayOfMonth = 31
End Function

Private Sub FILLEXPIRYGRID()
    Dim i As Integer
    Dim rstrefresh As ADODB.Recordset
    
    i = 0
    On Error GoTo eRRHAND
    Set rstrefresh = New ADODB.Recordset
    rstrefresh.Open "SELECT * from [EXPLIST] WHERE EX_FLAG ='N' ORDER BY EX_DISTI ", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstrefresh.EOF
        i = i + 1
        rstrefresh!EX_SLNO = i
        rstrefresh.MoveNext
    Loop
    rstrefresh.Close
    Set rstrefresh = Nothing
    
    Set rstrefresh = New ADODB.Recordset
    rstrefresh.Open "SELECT * from [EXPLIST] WHERE EX_FLAG ='Y' ORDER BY EX_DISTI ", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstrefresh.EOF
        i = i + 1
        rstrefresh!EX_SLNO = i
        rstrefresh.MoveNext
    Loop
    rstrefresh.Close
    Set rstrefresh = Nothing
    
    Screen.MousePointer = vbHourglass

    Set GrdEXPIRY.DataSource = Nothing
    If EXP_FLAG = True Then
        EXP_REC.Open "select * from [EXPLIST] WHERE EX_FLAG ='N' ORDER BY VAL(EX_SLNO)", db2, adOpenStatic, adLockReadOnly, adCmdText
        EXP_FLAG = False
    Else
        EXP_REC.Close
        EXP_REC.Open "select * from [EXPLIST] WHERE EX_FLAG ='N' ORDER BY VAL(EX_SLNO)", db2, adOpenStatic, adLockReadOnly, adCmdText
        EXP_FLAG = False
    End If
    
    
    Set GrdEXPIRY.DataSource = EXP_REC
    
    GrdEXPIRY.Columns(0).Caption = "SL"
    GrdEXPIRY.Columns(0).Width = 400
    GrdEXPIRY.Columns(1).Caption = "ITEM NAME"
    GrdEXPIRY.Columns(1).Width = 1900
    GrdEXPIRY.Columns(2).Visible = False
    GrdEXPIRY.Columns(3).Visible = False
    GrdEXPIRY.Columns(4).Caption = "COMPANY"
    GrdEXPIRY.Columns(4).Width = 1100
    GrdEXPIRY.Columns(5).Caption = "DISTRIBUTOR"
    GrdEXPIRY.Columns(5).Width = 900
    GrdEXPIRY.Columns(6).Caption = "BATCH"
    GrdEXPIRY.Columns(6).Width = 800
    GrdEXPIRY.Columns(7).Caption = "EXPIRY DATE"
    GrdEXPIRY.Columns(7).Width = 1150
    GrdEXPIRY.Columns(8).Caption = "QTY"
    GrdEXPIRY.Columns(8).Width = 500
    GrdEXPIRY.Columns(9).Caption = "MRP"
    GrdEXPIRY.Columns(9).Width = 500
    GrdEXPIRY.Columns(10).Visible = False
    GrdEXPIRY.Columns(11).Visible = False
    GrdEXPIRY.Columns(12).Visible = False
    GrdEXPIRY.Columns(13).Visible = False
      
    GrdEXPIRY.RowHeight = 250
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub FILLEXPIRYLIST()

    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    Set grdEXPIRYLIST.DataSource = Nothing
    If PHY_FLAG = True Then
        PHY_REC.Open "select * from [EXPIRY] ORDER BY EX_DATE,EX_ITEM", db2, adOpenStatic, adLockReadOnly, adCmdText
        PHY_FLAG = False
    Else
        PHY_REC.Close
        PHY_REC.Open "select * from [EXPIRY] ORDER BY EX_DATE,EX_ITEM", db2, adOpenStatic, adLockReadOnly, adCmdText
        PHY_FLAG = False
    End If
    
    
    Set grdEXPIRYLIST.DataSource = PHY_REC
    
    grdEXPIRYLIST.Columns(0).Visible = False
    grdEXPIRYLIST.Columns(1).Caption = "ITEM NAME"
    grdEXPIRYLIST.Columns(1).Width = 1900
    grdEXPIRYLIST.Columns(2).Visible = False
    grdEXPIRYLIST.Columns(3).Visible = False
    grdEXPIRYLIST.Columns(4).Visible = False
    'grdEXPIRYLIST.Columns(4).Caption = "COMPANY"
    'grdEXPIRYLIST.Columns(4).Width = 900
    'grdEXPIRYLIST.Columns(5).Visible = False
    grdEXPIRYLIST.Columns(5).Caption = "DISTRIBUTOR"
    grdEXPIRYLIST.Columns(5).Width = 1700
    grdEXPIRYLIST.Columns(6).Visible = False
    'grdEXPIRYLIST.Columns(6).Caption = "BATCH"
    'grdEXPIRYLIST.Columns(6).Width = 800
    grdEXPIRYLIST.Columns(7).Caption = "EXPIRY DATE"
    grdEXPIRYLIST.Columns(7).Width = 1150
    grdEXPIRYLIST.Columns(8).Caption = "QTY"
    grdEXPIRYLIST.Columns(8).Width = 400
    grdEXPIRYLIST.Columns(9).Visible = False
    grdEXPIRYLIST.Columns(10).Visible = False
    grdEXPIRYLIST.Columns(11).Visible = False
    grdEXPIRYLIST.Columns(12).Visible = False
    grdEXPIRYLIST.Columns(13).Visible = False
    grdEXPIRYLIST.Columns(14).Visible = False
    grdEXPIRYLIST.Columns(15).Visible = False
    grdEXPIRYLIST.Columns(16).Width = 1000
    grdEXPIRYLIST.Columns(16).Caption = "INVOICE"
    
    
    grdEXPIRYLIST.RowHeight = 250
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub TXTEXPDATE_GotFocus()
    TXTEXPDATE.SelStart = 0
    TXTEXPDATE.SelLength = Len(TXTEXPDATE.Text)
End Sub

Private Sub TXTEXPDATE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Not (IsDate(TXTEXPDATE.Text)) Then Exit Sub
            TXTMFGR.SetFocus
                        
    End Select
    
End Sub

Private Sub TXTEXPDATE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc("/")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTEXPQTY_GotFocus()
    TXTEXPQTY.SelStart = 0
    TXTEXPQTY.SelLength = Len(TXTEXPQTY.Text)
End Sub

Private Sub TXTEXPQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTEXPQTY.Text) = 0 Then Exit Sub
            TXTUNIT.SetFocus
                        
    End Select
End Sub

Private Sub TXTEXPQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTITEM_GotFocus()
    TXTITEM.SelStart = 0
    TXTITEM.SelLength = Len(TXTITEM.Text)
End Sub

Private Sub TXTITEM_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Trim(TXTITEM.Text) = "" Then Exit Sub
            TXTEXPQTY.SetFocus
                        
    End Select
End Sub

Private Sub TXTITEM_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTMFGR_GotFocus()
    TXTMFGR.SelStart = 0
    TXTMFGR.SelLength = Len(TXTMFGR.Text)
End Sub

Private Sub TXTMFGR_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            'If Trim(TXTMFGR.Text) = "" Then Exit Sub
            TxtMRP.SetFocus
                        
    End Select
End Sub

Private Sub TXTMFGR_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTMRP_GotFocus()
    TxtMRP.SelStart = 0
    TxtMRP.SelLength = Len(TxtMRP.Text)
End Sub

Private Sub TXTMRP_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If CMDEXPADD.Caption = "&SAVE" Then
                CMDEXPADD.SetFocus
            Else
                CMDEDITEXPLIST.SetFocus
            End If
                        
    End Select
End Sub

Private Sub TXTMRP_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub FILLCOMBO()
    On Error GoTo eRRHAND
    
    Screen.MousePointer = vbHourglass
    Set CMBexpdist.DataSource = Nothing
    If CMB_FLAG = True Then
        CMB_REC.Open "select Distinct EX_DISTI from [EXPLIST] WHERE EX_FLAG ='N' ORDER BY EX_DISTI", db2, adOpenStatic, adLockReadOnly, adCmdText
        CMB_FLAG = False
    Else
        CMB_REC.Close
        CMB_REC.Open "select Distinct EX_DISTI from [EXPLIST] WHERE EX_FLAG ='N' ORDER BY EX_DISTI", db2, adOpenStatic, adLockReadOnly, adCmdText
        CMB_FLAG = False
    End If
    
    Set Me.CMBexpdist.RowSource = CMB_REC
    CMBexpdist.ListField = "EX_DISTI"
    CMBexpdist.BoundColumn = "EX_DISTI"
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub TXTUNIT_GotFocus()
    TXTUNIT.SelStart = 0
    TXTUNIT.SelLength = Len(TXTUNIT.Text)
End Sub

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then Exit Sub
            CMBDISTI.SetFocus
                        
    End Select
End Sub

Private Sub TXTUNIT_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub FILLGRID()
    Dim i As Integer
    Dim rstrefresh As ADODB.Recordset
    
    i = 0
    On Error GoTo eRRHAND
    Screen.MousePointer = vbHourglass
    
    Set rstrefresh = New ADODB.Recordset
    rstrefresh.Open "SELECT * FROM EXPLIST where EXPLIST.EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db2, adOpenStatic, adLockOptimistic, adCmdText
    Do Until rstrefresh.EOF
        i = i + 1
        rstrefresh!EX_SLNO = i
        rstrefresh.MoveNext
    Loop
    rstrefresh.Close
    Set rstrefresh = Nothing

    Set GrdEXPIRY.DataSource = Nothing
    If EXP_FLAG = True Then
        EXP_REC.Open "SELECT * FROM EXPLIST where EXPLIST.EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db2, adOpenStatic, adLockReadOnly
        EXP_FLAG = False
    Else
        EXP_REC.Close
        EXP_REC.Open "SELECT * FROM EXPLIST where EXPLIST.EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db2, adOpenStatic, adLockReadOnly
        EXP_FLAG = False
    End If
    
    
    Set GrdEXPIRY.DataSource = EXP_REC
    
    GrdEXPIRY.Columns(0).Caption = "SL"
    GrdEXPIRY.Columns(0).Width = 400
    GrdEXPIRY.Columns(1).Caption = "ITEM NAME"
    GrdEXPIRY.Columns(1).Width = 1900
    GrdEXPIRY.Columns(2).Visible = False
    GrdEXPIRY.Columns(3).Visible = False
    GrdEXPIRY.Columns(4).Caption = "COMPANY"
    GrdEXPIRY.Columns(4).Width = 1100
    GrdEXPIRY.Columns(5).Caption = "DISTRIBUTOR"
    GrdEXPIRY.Columns(5).Width = 900
    GrdEXPIRY.Columns(6).Caption = "BATCH"
    GrdEXPIRY.Columns(6).Width = 800
    GrdEXPIRY.Columns(7).Caption = "EXPIRY DATE"
    GrdEXPIRY.Columns(7).Width = 1150
    GrdEXPIRY.Columns(8).Caption = "QTY"
    GrdEXPIRY.Columns(8).Width = 500
    GrdEXPIRY.Columns(9).Caption = "MRP"
    GrdEXPIRY.Columns(9).Width = 500
    GrdEXPIRY.Columns(10).Visible = False
    GrdEXPIRY.Columns(11).Visible = False
    GrdEXPIRY.Columns(12).Visible = False
    GrdEXPIRY.Columns(13).Visible = False
      
    GrdEXPIRY.RowHeight = 250
    Screen.MousePointer = vbNormal
    Exit Sub

eRRHAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub cmdReportGenerate_Click()
    Dim RSTEXP As ADODB.Recordset
    Dim vlineCount As Integer
    Dim vpageCount As Integer
    Dim TOTAL As Double
    Dim i As Integer
    
    vlineCount = 0
    vpageCount = 1
    SN = 0
    
    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
    '//NOTE : Report file name should never contain blank space.
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    
    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    Print #1, Chr(27) & Chr(72)
    Print #1, Chr(27) & Chr(71) & Chr(10) & Space(6) & " Name of Retailer:" & _
              Space(7) & Chr(14) & Chr(15) & "SARAS MEDICALS" & _
              Chr(27) & Chr(72)
    Print #1, Space(35) & " Kaichoondy, Alappuzha." & Chr(27) & Chr(67) & Chr(0) & Space(51) & Date

   ' Print #1, Space(7) & "Alappuzha 688006" & Space(15) & "DL No. 6-176/20/2003 Dtd. 31.10.2003"
    Print #1, Chr(27) & Chr(71) & Chr(10) & Space(6) & " Name of Distributor:" & _
              Space(4) & Chr(14) & Chr(15) & Trim(CMBexpdist.Text) & _
              Chr(27) & Chr(72)
    Print #1,
    
    Print #1, Space(9) & AlignLeft(" SL", 2) & Space(1) & _
            AlignLeft("ITEM NAME", 11) & Space(12) & _
            AlignLeft("INVOICE", 10) & Space(5) & _
            AlignLeft("INV DATE", 9) & Space(6) & _
            AlignLeft("MFGR", 15) & _
            AlignLeft("BATCH", 10) & _
            AlignLeft("EXP DATE", 11) & _
            AlignLeft("QTY", 7) & _
            AlignLeft("MRP", 8) & _
            AlignLeft("VALUE", 8) & _
            Chr(27) & Chr(72)  '//Bold Ends
            
    Print #1, Space(7) & RepeatString("-", 118)
    i = 0
    TOTAL = 0
    Set RSTEXP = New ADODB.Recordset
    RSTEXP.Open "SELECT * From EXPLIST WHERE EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db2, adOpenStatic, adLockReadOnly
    Do Until RSTEXP.EOF
            i = i + 1
         Print #1, Space(7) & AlignRight(Str(i), 3) & Space(2) & _
            AlignLeft(RSTEXP!EX_ITEM, 26) & _
            AlignLeft(RSTEXP!EX_PUR_INV, 10) & Space(1) & _
            AlignLeft(RSTEXP!EX_PUR_DATE, 15) & _
            AlignLeft(RSTEXP!EX_MFGR, 15) & _
            AlignLeft(RSTEXP!EX_BATCH, 12) & Space(1) & _
            AlignLeft(RSTEXP!EX_DATE, 7) & _
            AlignRight(RSTEXP!EX_QTY, 4) & Space(1) & _
            AlignRight(Format(RSTEXP!EX_MRP, ".00"), 8) & _
            AlignRight(Format((Val(RSTEXP!EX_MRP) * Val(RSTEXP!EX_QTY)) / Val(RSTEXP!EX_UNIT), ".00"), 9) & _
            Chr(27) & Chr(72)  '//Bold Ends
            TOTAL = TOTAL + ((Val(RSTEXP!EX_MRP) * Val(RSTEXP!EX_QTY)) / Val(RSTEXP!EX_UNIT))
        Print #1,
        RSTEXP.MoveNext
            
    Loop

    RSTEXP.Close
    Set RSTEXP = Nothing
    
    Print #1, Space(115) & AlignLeft("-------------", 10)
    Print #1, Space(102) & AlignLeft("NET AMOUNT", 10) & AlignRight((Format(TOTAL, "####.00")), 10)
    'Print #1, Chr(27) & Chr(71) & Chr(10) & Space(80) & AlignRight("NET AMOUNT", 10) & AlignRight((Format(TOTAL, "####.00")), 10)
    'Print #1, Chr(27) & Chr(67) & Chr(0)
   
    
    Close #1 '//Closing the file
    
End Sub

Private Sub EXPIRYReport()
    Dim RSTEXP As ADODB.Recordset
    Dim vlineCount As Integer
    Dim vpageCount As Integer
    Dim i As Integer
    
    vlineCount = 0
    vpageCount = 1
    SN = 0
    
    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
    '//NOTE : Report file name should never contain blank space.
    Open App.Path & "\Report.txt" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    
    Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    Print #1,
    
    Print #1, Space(9) & AlignLeft(" SL", 2) & Space(1) & _
            AlignLeft("ITEM NAME", 11) & Space(15) & _
            AlignLeft("BATCH", 11) & _
            AlignLeft("EXP DATE", 11) & _
            AlignLeft("QTY", 7) & _
            AlignLeft("MRP", 8) & _
            Chr(27) & Chr(72)  '//Bold Ends
            
    Print #1, Space(7) & RepeatString("-", 118)
    i = 0
 
    Set RSTEXP = New ADODB.Recordset
    RSTEXP.Open "SELECT * From EXPIRY ORDER BY EX_ITEM", db2, adOpenStatic, adLockReadOnly
    Do Until RSTEXP.EOF
            i = i + 1
         Print #1, Space(7) & AlignRight(Str(i), 3) & Space(2) & _
            AlignLeft(RSTEXP!EX_ITEM, 26) & _
            AlignLeft(RSTEXP!EX_BATCH, 12) & Space(1) & _
            AlignLeft(RSTEXP!EX_DATE, 7) & _
            AlignRight(RSTEXP!EX_QTY, 4) & Space(1) & _
            AlignRight(Format(RSTEXP!EX_MRP, ".00"), 8) & _
            Chr(27) & Chr(72)  '//Bold Ends
        Print #1,
        RSTEXP.MoveNext
            
    Loop

    RSTEXP.Close
    Set RSTEXP = Nothing
    
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
    Print #1, Chr(13)
    
    Close #1 '//Closing the file
    
End Sub

