VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FRMRETURN 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Return Warranty Products to the Customer"
   ClientHeight    =   8790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   17025
   ClipControls    =   0   'False
   Icon            =   "FRMEXPRETURN.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   17025
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   390
      Picture         =   "FRMEXPRETURN.frx":030A
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
      Picture         =   "FRMEXPRETURN.frx":064C
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
      Cols            =   15
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
      BackColor       =   &H00C0C0FF&
      Height          =   885
      Left            =   45
      TabIndex        =   0
      Top             =   7845
      Width           =   7260
      Begin VB.CommandButton Command2 
         Caption         =   "Move Items to Stock"
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
         Left            =   2715
         TabIndex        =   8
         Top             =   225
         Width           =   2580
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Return Items to Customers"
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
Attribute VB_Name = "FRMRETURN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim RSTFLAG As ADODB.Recordset
    Dim i, N As Integer
    
    If grdcount.Rows = 0 Then Exit Sub
    On Error GoTo Errhand
    
    N = 0
    If (MsgBox("ARE YOU SURE YOU WANT TO RETURN THESE ITEMS TO CUSTOMERS", vbYesNo) = vbNo) Then Exit Sub

    For i = 0 To grdcount.Rows - 1
        Set RSTFLAG = New ADODB.Recordset
        RSTFLAG.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdcount.TextMatrix(i, 0)) & " AND LINE_NO = " & Val(grdcount.TextMatrix(i, 5)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTFLAG.EOF And RSTFLAG.BOF) Then
            N = N + 1
            RSTFLAG!RETURN_FLAG = "Y"
            RSTFLAG!RETURN_DATE = Date
            RSTFLAG.Update
        End If
        RSTFLAG.Close
        Set RSTFLAG = Nothing
    Next i
    grdcount.Rows = 0
    Call Fillgrid
    MsgBox N & " Items Moved Successfully", vbOKOnly, "Warranty Replacement"
        
    Screen.MousePointer = vbNormal
    Exit Sub

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Command2_Click()
    Dim RSTFLAG As ADODB.Recordset
    Dim RSTTRXFILE As ADODB.Recordset
    Dim TRXMAST As ADODB.Recordset
    Dim i As Long
    Dim SLNO As Long
    
    If grdcount.Rows = 0 Then Exit Sub
    On Error GoTo Errhand
    
    If (MsgBox("ARE YOU SURE YOU WANT TO MOVE THESE ITEMS TO STOCK", vbYesNo) = vbNo) Then Exit Sub
    
    SLNO = 1
    Set TRXMAST = New ADODB.Recordset
    TRXMAST.Open "Select MAX(VCH_NO) From RTRXFILE WHERE TRX_TYPE = 'WR'", db, adOpenStatic, adLockReadOnly
    If Not (TRXMAST.EOF And TRXMAST.BOF) Then
        SLNO = IIf(IsNull(TRXMAST.Fields(0)), 1, TRXMAST.Fields(0) + 1)
    End If
    TRXMAST.Close
    Set TRXMAST = Nothing
    
    For i = 0 To grdcount.Rows - 1
        Set RSTFLAG = New ADODB.Recordset
        RSTFLAG.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(grdcount.TextMatrix(i, 0)) & " AND LINE_NO = " & Val(grdcount.TextMatrix(i, 5)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
        If Not (RSTFLAG.EOF And RSTFLAG.BOF) Then
            RSTFLAG!RETURN_FLAG = "S"
            RSTFLAG!RETURN_DATE = Date
            RSTFLAG.Update
        End If
        RSTFLAG.Close
        Set RSTFLAG = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdcount.TextMatrix(i, 6) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
        With RSTTRXFILE
            If Not (.EOF And .BOF) Then
                !ISSUE_QTY = !ISSUE_QTY - Val(grdcount.TextMatrix(i, 2))
                !CLOSE_QTY = !CLOSE_QTY + Val(grdcount.TextMatrix(i, 2))
                .Update
            End If
            .Close
        End With
        Set RSTTRXFILE = Nothing
        
        Set RSTTRXFILE = New ADODB.Recordset
        RSTTRXFILE.Open "SELECT *  FROM RTRXFILE ", db, adOpenStatic, adLockOptimistic, adCmdText
        
        RSTTRXFILE.AddNew
        RSTTRXFILE!TRX_TYPE = "WR"
        RSTTRXFILE!VCH_NO = SLNO
        RSTTRXFILE!line_no = i + 1
        RSTTRXFILE!ITEM_CODE = Trim(grdcount.TextMatrix(i, 6))
        RSTTRXFILE!QTY = Val(grdcount.TextMatrix(i, 2))
        RSTTRXFILE!BAL_QTY = Val(grdcount.TextMatrix(i, 2))
        RSTTRXFILE!TRX_TOTAL = 0
        RSTTRXFILE!VCH_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!ITEM_NAME = Trim(grdcount.TextMatrix(i, 1))
        RSTTRXFILE!ITEM_COST = 0
        RSTTRXFILE!LINE_DISC = 1
        RSTTRXFILE!P_DISC = 0
        RSTTRXFILE!MRP = 0
        RSTTRXFILE!PTR = 0
        RSTTRXFILE!SALES_PRICE = 0
        RSTTRXFILE!P_RETAIL = 0
        RSTTRXFILE!P_WS = 0
        RSTTRXFILE!P_CRTN = 0
        RSTTRXFILE!CRTN_PACK = 0
        RSTTRXFILE!P_VAN = 0
        RSTTRXFILE!COM_FLAG = ""
        RSTTRXFILE!COM_PER = 0
        RSTTRXFILE!COM_AMT = 0
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!UNIT = 1 'Val(grdcount.TextMatrix(N, 4))
        RSTTRXFILE!VCH_DESC = "Replaced From " & Trim(grdcount.TextMatrix(i, 4))
        RSTTRXFILE!REF_NO = Trim(grdcount.TextMatrix(i, 3))
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!SCHEME = 0
        'RSTTRXFILE!EXP_DATE = Null
        RSTTRXFILE!FREE_QTY = 0
        RSTTRXFILE!CREATE_DATE = Format(Date, "dd/mm/yyyy")
        RSTTRXFILE!C_USER_ID = "SM"
        RSTTRXFILE!M_USER_ID = "SM"
        RSTTRXFILE!CHECK_FLAG = "N"
        RSTTRXFILE!PINV = ""
        RSTTRXFILE.Update
    
        RSTTRXFILE.Close
        Set RSTTRXFILE = Nothing
        
    Next i
    grdcount.Rows = 0
    Call Fillgrid
    MsgBox "Items Moved Successfully", vbOKOnly, "Warranty Replacement"
        
    Screen.MousePointer = vbNormal
    Exit Sub

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub Form_Load()

    On Error GoTo Errhand
    
    Call Fillgrid
    
    Me.Width = 17145
    Me.Height = 9300
    Exit Sub
Errhand:
    MsgBox Err.Description
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.PCTMENU.Enabled = True
    'MDIMAIN.PCTMENU.Height = 15555
    MDIMAIN.PCTMENU.SetFocus
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
                    .TextMatrix(.Row, 13) = "Y"
                    Call fillcount
                Else
                    Set grdEXPIRYLIST.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 13) = "N"
                    Call fillcount
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Function fillcount()
    Dim i, N As Long
    
    grdcount.Rows = 0
    i = 0
    On Error GoTo Errhand
    For N = 1 To grdEXPIRYLIST.Rows - 1
        If grdEXPIRYLIST.TextMatrix(N, 13) = "Y" Then
            grdcount.Rows = grdcount.Rows + 1
            grdcount.TextMatrix(i, 0) = grdEXPIRYLIST.TextMatrix(N, 2)
            grdcount.TextMatrix(i, 1) = grdEXPIRYLIST.TextMatrix(N, 4)
            grdcount.TextMatrix(i, 2) = grdEXPIRYLIST.TextMatrix(N, 5)
            grdcount.TextMatrix(i, 3) = grdEXPIRYLIST.TextMatrix(N, 6)
            grdcount.TextMatrix(i, 4) = grdEXPIRYLIST.TextMatrix(N, 8)
            grdcount.TextMatrix(i, 5) = grdEXPIRYLIST.TextMatrix(N, 14)
            grdcount.TextMatrix(i, 6) = grdEXPIRYLIST.TextMatrix(N, 11)
            i = i + 1
        End If
    Next N
    Exit Function
Errhand:
    MsgBox Err.Description
    
End Function

Public Function Fillgrid()
    Dim RSTWARTRX As ADODB.Recordset
    Dim RSTRXFILE As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo Errhand

    Screen.MousePointer = vbHourglass
    
    i = 0
    grdEXPIRYLIST.FixedRows = 0
    grdEXPIRYLIST.Rows = 1
    
    grdEXPIRYLIST.TextMatrix(0, 0) = ""
    grdEXPIRYLIST.TextMatrix(0, 1) = "SL"
    grdEXPIRYLIST.TextMatrix(0, 2) = "REF NO"
    grdEXPIRYLIST.TextMatrix(0, 3) = "RET. DATE"
    grdEXPIRYLIST.TextMatrix(0, 4) = "ITEM NAME"
    grdEXPIRYLIST.TextMatrix(0, 5) = "QTY"
    grdEXPIRYLIST.TextMatrix(0, 6) = "SERIAL NO."
    grdEXPIRYLIST.TextMatrix(0, 7) = "CUSTOMER NAME"
    grdEXPIRYLIST.TextMatrix(0, 8) = "DISTRIBUTOR NAME"
    grdEXPIRYLIST.TextMatrix(0, 9) = "Inv. No"
    grdEXPIRYLIST.TextMatrix(0, 10) = "Inv Date"
    grdEXPIRYLIST.TextMatrix(0, 11) = "ITEM_CODE"
    grdEXPIRYLIST.TextMatrix(0, 12) = "CUSTOMER CODE"
    grdEXPIRYLIST.TextMatrix(0, 13) = "FLAG"
    grdEXPIRYLIST.TextMatrix(0, 14) = "LINE"
    
    grdEXPIRYLIST.ColWidth(0) = 300
    grdEXPIRYLIST.ColWidth(1) = 500
    grdEXPIRYLIST.ColWidth(2) = 1000
    grdEXPIRYLIST.ColWidth(3) = 1500
    grdEXPIRYLIST.ColWidth(4) = 2500
    grdEXPIRYLIST.ColWidth(5) = 800
    grdEXPIRYLIST.ColWidth(6) = 2400
    grdEXPIRYLIST.ColWidth(7) = 2600
    grdEXPIRYLIST.ColWidth(8) = 2600
    grdEXPIRYLIST.ColWidth(9) = 1200
    grdEXPIRYLIST.ColWidth(10) = 1200
    grdEXPIRYLIST.ColWidth(11) = 0
    grdEXPIRYLIST.ColWidth(12) = 0
    grdEXPIRYLIST.ColWidth(13) = 0
    grdEXPIRYLIST.ColWidth(14) = 0
    
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
        .Open "SELECT * From WAR_TRXFILE WHERE SETTLE_FLAG='Y' AND (ISNULL(RETURN_FLAG ) OR RETURN_FLAG ='N') ORDER BY VCH_NO", db, adOpenForwardOnly
        
        Do Until .EOF
            i = i + 1
            grdEXPIRYLIST.Rows = grdEXPIRYLIST.Rows + 1
            grdEXPIRYLIST.FixedRows = 1
            grdEXPIRYLIST.TextMatrix(i, 0) = ""
            grdEXPIRYLIST.TextMatrix(i, 1) = i
            grdEXPIRYLIST.TextMatrix(i, 2) = !VCH_NO
            grdEXPIRYLIST.TextMatrix(i, 3) = IIf(IsNull(!ARR_DATE), "", !ARR_DATE)
            grdEXPIRYLIST.TextMatrix(i, 4) = IIf(IsNull(!ITEM_NAME), "", !ITEM_NAME)
            grdEXPIRYLIST.TextMatrix(i, 5) = IIf(IsNull(!QTY), "", !QTY)
            grdEXPIRYLIST.TextMatrix(i, 6) = IIf(IsNull(!ARR_REF_NO), "", !ARR_REF_NO)
            grdEXPIRYLIST.TextMatrix(i, 7) = IIf(IsNull(!ACT_NAME), "", !ACT_NAME)
            grdEXPIRYLIST.TextMatrix(i, 8) = IIf(IsNull(!DIST_NAME), "", !DIST_NAME)
            grdEXPIRYLIST.TextMatrix(i, 9) = IIf(IsNull(!BILL_NO), "", !BILL_NO)
            grdEXPIRYLIST.TextMatrix(i, 10) = IIf(IsNull(!BILL_DATE), "", !BILL_DATE)
            grdEXPIRYLIST.TextMatrix(i, 11) = IIf(IsNull(!ITEM_CODE), "", !ITEM_CODE)
            grdEXPIRYLIST.TextMatrix(i, 12) = IIf(IsNull(!ACT_CODE), "", !ACT_CODE)
            grdEXPIRYLIST.TextMatrix(i, 13) = "N"
            grdEXPIRYLIST.TextMatrix(i, 14) = !line_no
            With grdEXPIRYLIST
              .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
              Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
              .TextMatrix(i, 1) = i
            End With

            
            .MoveNext
        Loop
        .Close
        Set RSTWARTRX = Nothing
    End With
    
    Screen.MousePointer = vbNormal
    Exit Function

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function


