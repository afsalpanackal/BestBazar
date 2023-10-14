VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMEXPRCVDWO 
   BackColor       =   &H00FFFF80&
   Caption         =   "Receive Warranty Items from Supplier"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FRMEXPRCVDWO.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   17025
   ShowInTaskbar   =   0   'False
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
      Left            =   900
      TabIndex        =   9
      Top             =   3000
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSMask.MaskEdBox TXTEXPIRY 
      Height          =   285
      Left            =   8070
      TabIndex        =   8
      Top             =   2205
      Visible         =   0   'False
      Width           =   1560
      _ExtentX        =   2752
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/####"
      PromptChar      =   " "
   End
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   390
      Picture         =   "FRMEXPRCVDWO.frx":030A
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1140
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   720
      Picture         =   "FRMEXPRCVDWO.frx":064C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   1170
      Visible         =   0   'False
      Width           =   285
   End
   Begin MSFlexGridLib.MSFlexGrid grdEXPIRYLIST 
      Height          =   7590
      Left            =   15
      TabIndex        =   3
      Top             =   255
      Width           =   16965
      _ExtentX        =   29924
      _ExtentY        =   13388
      _Version        =   393216
      Rows            =   1
      Cols            =   18
      FixedRows       =   0
      RowHeightMin    =   400
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
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame FRMECONTROLS 
      BackColor       =   &H00FFFF80&
      Height          =   885
      Left            =   45
      TabIndex        =   0
      Top             =   7845
      Width           =   7260
      Begin VB.CommandButton Command1 
         Caption         =   "Items Settled and moved to Warehouse"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   45
         TabIndex        =   6
         Top             =   225
         Width           =   2595
      End
      Begin VB.CommandButton CMDEXIT 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   5400
         TabIndex        =   1
         Top             =   225
         Width           =   1740
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdcount 
      Height          =   5145
      Left            =   -15
      TabIndex        =   7
      Top             =   750
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   9075
      _Version        =   393216
      Rows            =   1
      Cols            =   15
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
   Begin MSForms.CommandButton CommandButton1 
      Height          =   30
      Left            =   1185
      TabIndex        =   2
      Top             =   4035
      Width           =   45
      Size            =   "79;53"
      FontHeight      =   165
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
   End
End
Attribute VB_Name = "FRMEXPRCVDWO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CLOSEALL As Integer
Dim M_STOCK As Integer
Dim M_EDIT As Boolean
Dim NONSTOCK As Boolean
'Dim strChecked As String

Private Sub CMDEXIT_Click()
    CLOSEALL = 0
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim RSTFLAG As ADODB.Recordset
    Dim I, n As Integer
    
    If grdcount.Rows = 0 Then Exit Sub
    On Error GoTo eRRhAND
    
    n = 0
    If (MsgBox("ARE YOU SURE YOU WANT TO MOVE THESE ITEMS TO WAREHOUSE", vbYesNo) = vbNo) Then Exit Sub

    For I = 0 To grdcount.Rows - 1
        Set RSTFLAG = New ADODB.Recordset
        RSTFLAG.Open "SELECT * from [WAR_TRXFILE] where VCH_NO = " & Val(grdcount.TextMatrix(I, 0)) & " AND LINE_NO = " & Val(grdcount.TextMatrix(I, 1)) & " ", db2, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTFLAG.EOF And RSTFLAG.BOF) Then
            n = n + 1
            RSTFLAG!SETTLE_FLAG = "Y"
            RSTFLAG!SETTLE_DATE = Date
            RSTFLAG!ARR_REF_NO = Trim(grdcount.TextMatrix(I, 7))
            RSTFLAG!ARR_BILL_NO = Trim(grdcount.TextMatrix(I, 8))
            RSTFLAG!ARR_DATE = IIf(IsDate(grdcount.TextMatrix(I, 9)), grdcount.TextMatrix(I, 9), Null)
            RSTFLAG.Update
        End If
        RSTFLAG.Close
        Set RSTFLAG = Nothing
    Next I
    grdcount.Rows = 0
    Call FILLGRID
    MsgBox n & " Items Moved Successfully", vbOKOnly, "Warranty Replacement"
        
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Load()

    On Error GoTo eRRhAND
    
    Call FILLGRID
    
    Me.Width = 17145
    Me.Height = 9300
    Exit Sub
eRRhAND:
    MsgBox Err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If CLOSEALL = 0 Then
        'ACT_REC.Close
        'MFG_REC.Close
    
        MDIMAIN.PCTMENU.Enabled = True
        'MDIMAIN.PCTMENU.Height = 15555
        MDIMAIN.PCTMENU.SetFocus
    End If
    Cancel = CLOSEALL
End Sub

Private Sub grdEXPIRYLIST_Click()
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If grdEXPIRYLIST.Rows = 1 Then Exit Sub
    If grdEXPIRYLIST.Col <> 1 Then Exit Sub
    With grdEXPIRYLIST
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If grdEXPIRYLIST.Col = 0 Then
                If grdEXPIRYLIST.CellPicture = picChecked Then
                    Set grdEXPIRYLIST.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 16) = "Y"
                    Call fillcount
                Else
                    Set grdEXPIRYLIST.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 16) = "N"
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Function fillcount()
    Dim I, n As Long
    
    grdcount.Rows = 0
    I = 0
    On Error GoTo eRRhAND
    For n = 1 To grdEXPIRYLIST.Rows - 1
        If grdEXPIRYLIST.TextMatrix(n, 16) = "Y" Then
            grdcount.Rows = grdcount.Rows + 1
            grdcount.TextMatrix(I, 0) = grdEXPIRYLIST.TextMatrix(n, 2)
            grdcount.TextMatrix(I, 1) = grdEXPIRYLIST.TextMatrix(n, 17)

            I = I + 1
        End If
    Next n
    Exit Function
eRRhAND:
    MsgBox Err.Description
    
End Function

Public Function FILLGRID()
    Dim RSTWARTRX As ADODB.Recordset
    Dim RSTRXFILE As ADODB.Recordset
    Dim I As Integer
    
    On Error GoTo eRRhAND

    Screen.MousePointer = vbHourglass
    
    I = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.Rows = 1
    
    grdEXPIRYLIST.TextMatrix(0, 0) = ""
    grdEXPIRYLIST.TextMatrix(0, 1) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 2) = "REF NO"
    grdEXPIRYLIST.TextMatrix(0, 3) = "DATE"
    grdEXPIRYLIST.TextMatrix(0, 4) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 5) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 6) = "SERIAL NO."
    grdEXPIRYLIST.TextMatrix(0, 7) = "Arr. SERIAL NO."
    grdEXPIRYLIST.TextMatrix(0, 8) = "Arr. Ref No."
    grdEXPIRYLIST.TextMatrix(0, 9) = "Arr. Date."
    grdEXPIRYLIST.TextMatrix(0, 10) = "CUSTOMER NAME"
    grdEXPIRYLIST.TextMatrix(0, 11) = "DISTRIBUTOR NAME"
    grdEXPIRYLIST.TextMatrix(0, 12) = "Inv. No"
    grdEXPIRYLIST.TextMatrix(0, 13) = "Inv Date"
    grdEXPIRYLIST.TextMatrix(0, 14) = "ITEM_CODE"
    grdEXPIRYLIST.TextMatrix(0, 15) = "CUSTOMER CODE"
    grdEXPIRYLIST.TextMatrix(0, 16) = "FLAG"
    grdEXPIRYLIST.TextMatrix(0, 17) = "LINE"
    
    grdEXPIRYLIST.ColWidth(0) = 300
    grdEXPIRYLIST.ColWidth(1) = 500
    grdEXPIRYLIST.ColWidth(2) = 1000
    grdEXPIRYLIST.ColWidth(3) = 1500
    grdEXPIRYLIST.ColWidth(4) = 2500
    grdEXPIRYLIST.ColWidth(5) = 800
    grdEXPIRYLIST.ColWidth(6) = 2400
    grdEXPIRYLIST.ColWidth(7) = 2400
    grdEXPIRYLIST.ColWidth(8) = 1800
    grdEXPIRYLIST.ColWidth(9) = 1400
    grdEXPIRYLIST.ColWidth(10) = 2600
    grdEXPIRYLIST.ColWidth(11) = 2600
    grdEXPIRYLIST.ColWidth(12) = 1200
    grdEXPIRYLIST.ColWidth(13) = 1200
    grdEXPIRYLIST.ColWidth(14) = 0
    grdEXPIRYLIST.ColWidth(15) = 0
    grdEXPIRYLIST.ColWidth(16) = 0
    grdEXPIRYLIST.ColWidth(17) = 0
    
    grdEXPIRYLIST.ColAlignment(0) = 9
    grdEXPIRYLIST.ColAlignment(1) = 4
    grdEXPIRYLIST.ColAlignment(2) = 4
    grdEXPIRYLIST.ColAlignment(3) = 4
    grdEXPIRYLIST.ColAlignment(4) = 9
    grdEXPIRYLIST.ColAlignment(5) = 4
    grdEXPIRYLIST.ColAlignment(6) = 1
    grdEXPIRYLIST.ColAlignment(7) = 1
    grdEXPIRYLIST.ColAlignment(8) = 1
    grdEXPIRYLIST.ColAlignment(9) = 4
    grdEXPIRYLIST.ColAlignment(10) = 4
    
    Set RSTWARTRX = New ADODB.Recordset
    With RSTWARTRX
        .Open "SELECT * From [WAR_TRXFILE] WHERE CHECK_FLAG='Y' AND (ISNULL(SETTLE_FLAG ) OR SETTLE_FLAG <>'Y') ORDER BY VCH_NO", db2, adOpenForwardOnly
        
        Do Until .EOF
            I = I + 1
            grdEXPIRYLIST.Rows = grdEXPIRYLIST.Rows + 1
            grdEXPIRYLIST.FixedRows = 1
            grdEXPIRYLIST.TextMatrix(I, 0) = ""
            grdEXPIRYLIST.TextMatrix(I, 1) = I
            grdEXPIRYLIST.TextMatrix(I, 2) = !VCH_NO
            grdEXPIRYLIST.TextMatrix(I, 3) = !VCH_DATE
            grdEXPIRYLIST.TextMatrix(I, 4) = IIf(IsNull(!ITEM_NAME), "", !ITEM_NAME)
            grdEXPIRYLIST.TextMatrix(I, 5) = IIf(IsNull(!QTY), "", !QTY)
            grdEXPIRYLIST.TextMatrix(I, 6) = IIf(IsNull(!REF_NO), "", !REF_NO)
            grdEXPIRYLIST.TextMatrix(I, 7) = ""
            grdEXPIRYLIST.TextMatrix(I, 8) = ""
            grdEXPIRYLIST.TextMatrix(I, 9) = ""
            grdEXPIRYLIST.TextMatrix(I, 10) = IIf(IsNull(!ACT_NAME), "", !ACT_NAME)
            grdEXPIRYLIST.TextMatrix(I, 11) = IIf(IsNull(!DIST_NAME), "", !DIST_NAME)
            grdEXPIRYLIST.TextMatrix(I, 12) = IIf(IsNull(!BILL_NO), "", !BILL_NO)
            grdEXPIRYLIST.TextMatrix(I, 13) = IIf(IsNull(!BILL_DATE), "", !BILL_DATE)
            grdEXPIRYLIST.TextMatrix(I, 14) = IIf(IsNull(!ITEM_CODE), "", !ITEM_CODE)
            grdEXPIRYLIST.TextMatrix(I, 15) = IIf(IsNull(!ACT_CODE), "", !ACT_CODE)
            grdEXPIRYLIST.TextMatrix(I, 16) = "N"
            grdEXPIRYLIST.TextMatrix(I, 17) = !LINE_NO
            With grdEXPIRYLIST
              .Row = I: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(I, 1) = I
            End With

            
            .MoveNext
        Loop
        .Close
        Set RSTWARTRX = Nothing
    End With
    
    Screen.MousePointer = vbNormal
    Exit Function

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function


Private Sub grdEXPIRYLIST_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeySpace
            Call grdEXPIRYLIST_Click
        Case vbKeyReturn
            Select Case grdEXPIRYLIST.Col
                Case 7, 8
                    Select Case grdEXPIRYLIST.Col
                        Case 9 'bILL nO
                            TXTsample.MaxLength = 15
                        Case 5 'qty
                            TXTsample.MaxLength = 6
                        Case 6 'batch
                            TXTsample.MaxLength = 30
                    End Select
                    TXTsample.Visible = True
                    TXTsample.Top = grdEXPIRYLIST.CellTop + 125
                    TXTsample.Left = grdEXPIRYLIST.CellLeft + 25
                    TXTsample.Width = grdEXPIRYLIST.CellWidth + 50
                    TXTsample.Text = grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col)
                    TXTsample.SetFocus
        
                Case 9
                    TXTEXPIRY.Visible = True
                    TXTEXPIRY.Top = grdEXPIRYLIST.CellTop + 325
                    TXTEXPIRY.Left = grdEXPIRYLIST.CellLeft + 75
                    TXTEXPIRY.Width = grdEXPIRYLIST.CellWidth
                    TXTEXPIRY.Height = grdEXPIRYLIST.CellHeight
                    If Not (IsDate(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col))) Then
                        TXTEXPIRY.Text = "  /  /    "
                    Else
                        TXTEXPIRY.Text = Format(grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col), "DD/MM/YYYY")
                    End If
                    
                    TXTEXPIRY.SetFocus
            End Select
    End Select

End Sub

Private Sub TXTEXPIRY_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
                If Not (IsDate(TXTEXPIRY.Text)) Then Exit Sub
                If Len(TXTEXPIRY.Text) < 10 Then Exit Sub
                If DateValue(TXTEXPIRY.Text) > DateValue(Date) Then
                    MsgBox "From Address could not be higher than Today", vbOKOnly, "IMEI REQUEST..."
                    TXTEXPIRY.SetFocus
                    Exit Sub
                End If
                
                grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = Format(TXTEXPIRY.Text, "dd/mm/yyyy")
                
                grdEXPIRYLIST.Enabled = True
                TXTEXPIRY.Visible = False
                grdEXPIRYLIST.SetFocus
        Case vbKeyEscape
            TXTEXPIRY.Visible = False
            grdEXPIRYLIST.SetFocus
    End Select
    Exit Sub
eRRhAND:
    MsgBox Err.Description

End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        Case vbKeyReturn
            Select Case grdEXPIRYLIST.Col
                Case 7   'batch
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
                    grdEXPIRYLIST.Enabled = True
                    TXTsample.Visible = False
                    grdEXPIRYLIST.SetFocus
                Case 8  'INV NO
                    'If Val(TXTsample.Text) = 0 Then Exit Sub
                    grdEXPIRYLIST.TextMatrix(grdEXPIRYLIST.Row, grdEXPIRYLIST.Col) = TXTsample.Text
                    grdEXPIRYLIST.Enabled = True
                    TXTsample.Visible = False
                    grdEXPIRYLIST.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            grdEXPIRYLIST.SetFocus
    End Select
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case grdEXPIRYLIST.Col
        Case 5, 4
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 1, 2, 6
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 8
             Select Case KeyAscii
                Case Asc("'")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub TXTsample_LostFocus()
    TXTsample.Visible = False
End Sub

Private Sub TXTEXPIRY_GotFocus()
    TXTEXPIRY.SelStart = 0
    TXTEXPIRY.SelLength = Len(TXTEXPIRY.Text)
End Sub

Private Sub TXTEXPIRY_LostFocus()
    TXTEXPIRY.Visible = False
End Sub

