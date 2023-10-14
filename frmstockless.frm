VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmstockless 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "STOCK LESS"
   ClientHeight    =   9750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmstockless.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   9810
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
      Height          =   5430
      Left            =   5775
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   2625
      Width           =   3915
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Height          =   945
      Left            =   5775
      TabIndex        =   13
      Top             =   8235
      Width           =   3870
      Begin VB.CommandButton CMDREMOVE 
         BackColor       =   &H00400000&
         Caption         =   "RE&MOVE DISTRIBUTOR"
         Height          =   495
         Left            =   2385
         MaskColor       =   &H80000007&
         TabIndex        =   9
         Top             =   225
         UseMaskColor    =   -1  'True
         Width           =   1380
      End
      Begin VB.TextBox txtunit 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   75
         MaxLength       =   3
         TabIndex        =   7
         Top             =   465
         Width           =   810
      End
      Begin VB.CommandButton CMDADDDIST 
         BackColor       =   &H00400000&
         Caption         =   "ADD DIS&TRIBUTOR"
         Height          =   495
         Left            =   960
         MaskColor       =   &H80000007&
         TabIndex        =   8
         Top             =   240
         UseMaskColor    =   -1  'True
         Width           =   1365
      End
      Begin VB.Label LBLUNI 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
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
         Left            =   45
         TabIndex        =   14
         Top             =   150
         Width           =   840
      End
   End
   Begin VB.ListBox LSTDUMMY 
      Height          =   1425
      Left            =   7230
      TabIndex        =   12
      Top             =   3135
      Visible         =   0   'False
      Width           =   1395
   End
   Begin VB.Frame frmeparty 
      Height          =   2625
      Left            =   5775
      TabIndex        =   10
      Top             =   -105
      Width           =   3915
      Begin VB.CommandButton CMDMINQTY 
         Caption         =   "MIN QTY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2655
         TabIndex        =   16
         Top             =   2145
         Width           =   1200
      End
      Begin VB.CommandButton Cmdexit 
         Caption         =   "E&XIT"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1410
         TabIndex        =   5
         Top             =   2145
         Width           =   1200
      End
      Begin VB.CommandButton CMDADD 
         Caption         =   "&Add Item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         TabIndex        =   4
         Top             =   2160
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
         Height          =   315
         Left            =   600
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1755
         Width           =   720
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
         ItemData        =   "frmstockless.frx":030A
         Left            =   1410
         List            =   "frmstockless.frx":0317
         TabIndex        =   3
         Top             =   1740
         Width           =   1260
      End
      Begin MSDataListLib.DataList DatalstSupplier 
         Height          =   1425
         Left            =   75
         TabIndex        =   1
         Top             =   285
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2514
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
         Left            =   90
         TabIndex        =   15
         Top             =   75
         Width           =   3690
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   165
         TabIndex        =   11
         Top             =   1755
         Width           =   405
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdSTOCKLESS 
      Height          =   9180
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   5730
      _ExtentX        =   10107
      _ExtentY        =   16193
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   300
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      AllowBigSelection=   0   'False
      FocusRect       =   2
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
End
Attribute VB_Name = "frmstockless"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TMPFLAG As Boolean
Dim rstTMP As New ADODB.Recordset
Dim M_STOCK As Integer

Private Sub CMDADD_Click()
    Dim i As Integer
    Dim SN As Integer
    Dim FCODE, FCODE1, FCODE2, FCODE3, FCODE4, FCODE5, FCODE6, FCODE7 As String
    'FCODE - CODE+CODE  FCODE1 - ITEM , FCODE2 - QTY, FCODE3 - DISTRIBUTOR
    'FCODE4 - ITEMCODE  FCODE5 - DISTCODE , FCODE6 - Or_Stock, FCODE7 - Or_AcQty
    
    If grdSTOCKLESS.Rows <= 1 Then Exit Sub

    If Trim(TXTQTY.Text) = "" Then
        MsgBox "Enter the Quantity", vbOKOnly, "ORDER"
        TXTQTY.SetFocus
        Exit Sub
    End If
    
    If DatalstSupplier.Text = "" Then
        MsgBox "SELECT DISTRIBUTOR", vbOKOnly, "ORDER"
        DatalstSupplier.SetFocus
        Exit Sub
    End If

    On Error GoTo eRRhAND
    
    FCODE7 = Val(TXTQTY.Text) * Val(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 4)) ' Actual Qty
    FCODE6 = Val(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 3)) ' AVL STOCK
    FCODE3 = Me.DatalstSupplier.Text 'DISTRIBUTOR
    FCODE1 = Trim(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 2)) 'ITEM
    If Val(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 4)) > 1 Then
        TXTQTY.Tag = " x " & Val(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 4)) & " "
    Else
        TXTQTY.Tag = " "
    End If
        
    If Val(TXTQTY) <= 1 Then
        FCODE2 = Me.TXTQTY.Text & TXTQTY.Tag & CmbQty.Text   'QTY
        
    Else
        If Right(CmbQty.Text, 1) = "s" Or Right(CmbQty.Text, 1) = "S" Or Trim(CmbQty.Text) = "" Then
            FCODE2 = Me.TXTQTY.Text & TXTQTY.Tag & CmbQty.Text   'QTY
        Else
            FCODE2 = Me.TXTQTY.Text & TXTQTY.Tag & CmbQty.Text & "s"   'QTY
        End If
    End If
    FCODE = Trim(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1)) & Trim(DatalstSupplier.BoundText)
    FCODE4 = Trim(grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1))
    FCODE5 = Trim(DatalstSupplier.BoundText)
    
    db2.Execute ("insert into TmpOrderlist values ('" & FCODE & "','" & FCODE1 & "','" & FCODE2 & "','" & FCODE3 & "','" & FCODE4 & "','" & FCODE5 & "','" & FCODE6 & "','" & FCODE7 & "' )")
    'SN = grdSTOCKLESS.Row
    'For i = SN To grdSTOCKLESS.Rows - 2
    '    grdSTOCKLESS.TextMatrix(SN, 0) = grdSTOCKLESS.TextMatrix(i, 0)
    '    grdSTOCKLESS.TextMatrix(SN, 1) = grdSTOCKLESS.TextMatrix(i + 1, 1)
    '    grdSTOCKLESS.TextMatrix(SN, 2) = grdSTOCKLESS.TextMatrix(i + 1, 2)
    '    grdSTOCKLESS.TextMatrix(SN, 3) = grdSTOCKLESS.TextMatrix(i + 1, 3)
    '    grdSTOCKLESS.TextMatrix(SN, 4) = grdSTOCKLESS.TextMatrix(i + 1, 4)
    '    SN = SN + 1
    'Next i
    'grdSTOCKLESS.Rows = grdSTOCKLESS.Rows - 1
       
    Call grdSTOCKLESS_Click
    TXTQTY.Text = ""
    CmbQty.Text = ""
   Exit Sub
   
eRRhAND:
    MsgBox "ALREADY ENTERED", vbOKOnly, "ORDER"
End Sub

Private Sub CMDADDDIST_Click()
    Dim RSTA As ADODB.Recordset
    Dim RSTB As ADODB.Recordset
    Dim RSTC As ADODB.Recordset
    
    Dim i As Integer
    
    If grdSTOCKLESS.Rows <= 1 Then Exit Sub
    
    If lstmanufact.SelCount = 0 Then
        MsgBox "Please Select the Distributor to be added", vbOKOnly, "ORDER"
        Exit Sub
    End If
    
    If TXTUNIT.Visible = True And Val(TXTUNIT.Text) = 0 Then
        MsgBox "Please Enter the Unit", vbOKOnly, "ORDER"
        TXTUNIT.SetFocus
        Exit Sub
    End If
    
    On Error GoTo eRRhAND
    
    i = 0
    
    Set RSTA = New ADODB.Recordset
    
    RSTA.Open "SELECT *  FROM PRODLINK WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    Set RSTB = New ADODB.Recordset
    RSTB.Open "SELECT *  FROM PRODLINK WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
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
                    !ACT_CODE = Mid(lstmanufact.List(i), 1, 6)
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
            RSTC.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
            For i = 0 To lstmanufact.ListCount - 1
                If lstmanufact.Selected(i) Then
                    .AddNew
                    !ITEM_CODE = RSTC!ITEM_CODE
                    !ITEM_NAME = RSTC!ITEM_NAME
                    !RQTY = Null
                    !ITEM_COST = RSTC!ITEM_COST
                    !MRP = RSTC!MRP
                    !PTR = RSTC!PTR
                    !SALES_PRICE = Val(RSTC!MRP) + (Val(RSTC!MRP) * Val(RSTC!SALES_TAX) / 100)
                    !SALES_TAX = RSTC!SALES_TAX
                    !UNIT = Val(TXTUNIT.Text)
                    !Remarks = Val(TXTUNIT.Text)
                    !ORD_QTY = 0
                    !CST = RSTC!CST
                    !ACT_CODE = Mid(lstmanufact.List(i), 1, 6)
                    !CREATE_DATE = Date
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
    

    TXTUNIT.Text = ""
    grdSTOCKLESS_Click
    DatalstSupplier.SetFocus
       
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub CMDMINQTY_Click()
    Me.Enabled = False
    frmminstock.Show
End Sub

Private Sub CMDREMOVE_Click()
    If grdSTOCKLESS.Rows <= 1 Then Exit Sub
    
    If DatalstSupplier.Text = "" Then
        Exit Sub
    End If

    If MsgBox("ARE YOU SURE YOU WANT TO REMOVE " & DatalstSupplier.Text, vbYesNo, "DELETE....?") = vbNo Then Exit Sub
    On Error GoTo eRRhAND
      
    db.Execute ("DELETE *  FROM PRODLINK WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "' AND ACT_CODE = '" & DatalstSupplier.BoundText & "'")
    grdSTOCKLESS_Click
    DatalstSupplier.SetFocus
       
    Exit Sub
    
eRRhAND:
    MsgBox Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If TMPFLAG = False Then rstTMP.Close
    FrmCrimedata.Enabled = True
    'MDIMAIN.PCTMENU.Enabled = False
    'MDIMAIN.PCTMENU.Height = 555
    FrmCrimedata.SetFocus
End Sub

Private Sub grdSTOCKLESS_Click()
    Dim RSTAVL As ADODB.Recordset
    Dim RSTMAN As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    Dim i As Integer
    Dim N As Integer
    
    On Error GoTo eRRhAND
    
    LBLMANUFACT.Caption = ""
    LSTDUMMY.Clear
    lstmanufact.Clear

    i = 0
    If grdSTOCKLESS.Rows <= 1 Then Exit Sub
    
    Set RSTAVL = New ADODB.Recordset
    RSTAVL.Open "SELECT MANUFACTURER, ITEM_NAME, ITEM_CODE FROM ITEMMAST WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTAVL
        LBLMANUFACT.Caption = RSTAVL!MANUFACTURER
        If Not (.EOF And .BOF) Then
            If TMPFLAG = True Then
                rstTMP.Open "SELECT PRODLINK.UNIT, PRODLINK.ITEM_CODE, PRODLINK.ITEM_NAME, ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST RIGHT JOIN PRODLINK ON ACTMAST.ACT_CODE = PRODLINK.ACT_CODE WHERE ITEM_CODE = '" & RSTAVL!ITEM_CODE & "' ORDER BY ACTMAST.ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                TMPFLAG = False
            Else
                rstTMP.Close
                rstTMP.Open "SELECT PRODLINK.UNIT, PRODLINK.ITEM_CODE, PRODLINK.ITEM_NAME, ACTMAST.ACT_CODE, ACTMAST.ACT_NAME FROM ACTMAST RIGHT JOIN PRODLINK ON ACTMAST.ACT_CODE = PRODLINK.ACT_CODE WHERE ITEM_CODE = '" & RSTAVL!ITEM_CODE & "' ORDER BY ACTMAST.ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
                TMPFLAG = False
            End If
        
            Set DatalstSupplier.RowSource = rstTMP
            DatalstSupplier.ListField = "ACT_NAME"
            DatalstSupplier.BoundColumn = "ACT_CODE"
            
            i = 0
            Do Until rstTMP.EOF
            
                LSTDUMMY.AddItem (i)
                LSTDUMMY.List(i) = rstTMP!ACT_CODE
                i = i + 1
                rstTMP.MoveNext
            
            Loop
              
    End If
        .Close
    End With
    Set RSTAVL = Nothing
    
    i = 0
    
    Set RSTMAN = New ADODB.Recordset
    RSTMAN.Open "select ACT_CODE, ACT_NAME from [ACTMAST]  WHERE (Mid(ACT_CODE, 1, 3)='311')And (len(ACT_CODE)>3) ORDER BY ACT_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    With RSTMAN
        Do Until .EOF
            
            For N = 0 To LSTDUMMY.ListCount
                If Trim(LSTDUMMY.List(N)) = Trim(!ACT_CODE) Then GoTo SKIP
            Next N
            lstmanufact.AddItem (i)
            lstmanufact.List(i) = !ACT_CODE & " " & Trim(!ACT_NAME)
            i = i + 1
SKIP:
        .MoveNext
        Loop
        .Close
    End With
    Set RSTMAN = Nothing
    
    If DatalstSupplier.VisibleCount = 0 Then
        TXTUNIT.Visible = True
        LBLUNI.Visible = True
    Else
        TXTUNIT.Text = ""
        TXTUNIT.Visible = False
        LBLUNI.Visible = False
    End If
    
    Exit Sub
    
eRRhAND:
    If Err.Number = 3021 Then
        Resume Next
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub CMDEXIT_Click()
    MDIMAIN.PCTMENU.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim RSTSTKLESS As ADODB.Recordset
    Dim RSTORDER As ADODB.Recordset
    
    Dim i As Integer
    
    PHY_FLAG = True
    
    On Error GoTo eRRhAND

    
    i = 0
    grdSTOCKLESS.TextMatrix(0, 0) = "SL"
    grdSTOCKLESS.TextMatrix(0, 1) = "ITEM CODE"
    grdSTOCKLESS.TextMatrix(0, 2) = "ITEM NAME"
    grdSTOCKLESS.TextMatrix(0, 3) = "STOCK"
    grdSTOCKLESS.TextMatrix(0, 4) = "UNIT"
    grdSTOCKLESS.TextMatrix(0, 5) = "MIN STOCK"
    
    grdSTOCKLESS.ColWidth(0) = 500
    grdSTOCKLESS.ColWidth(1) = 0
    grdSTOCKLESS.ColWidth(2) = 3300
    grdSTOCKLESS.ColWidth(3) = 800
    grdSTOCKLESS.ColWidth(4) = 0
    grdSTOCKLESS.ColWidth(5) = 1000
    
    grdSTOCKLESS.ColAlignment(0) = 1
    grdSTOCKLESS.ColAlignment(3) = 3
    grdSTOCKLESS.ColAlignment(5) = 3
    Screen.MousePointer = vbHourglass
    
    Set RSTSTKLESS = New ADODB.Recordset
    'RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE WHERE  ITEMMAST.REORDER_QTY > ITEMMAST.CLOSE_QTY ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    RSTSTKLESS.Open "SELECT DISTINCT RTRXFILE.ITEM_CODE, RTRXFILE.ITEM_NAME, RTRXFILE.UNIT, ITEMMAST.ITEM_CODE, ITEMMAST.REORDER_QTY, ITEMMAST.CLOSE_QTY, ITEMMAST.ITEM_NAME FROM RTRXFILE RIGHT JOIN ITEMMAST ON RTRXFILE.ITEM_CODE = ITEMMAST.ITEM_CODE ORDER BY ITEMMAST.ITEM_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until RSTSTKLESS.EOF
        Set RSTORDER = New ADODB.Recordset
        RSTORDER.Open "Select DISTINCT item_Code, Or_Product From TMPORDERLIST WHERE [item_Code] ='" & RSTSTKLESS!ITEM_CODE & "'", db2, adOpenStatic, adLockReadOnly, adCmdText
        If RSTORDER.RecordCount > 0 Then GoTo SKIP
            'If IsNull(RSTSTKLESS!CLOSE_QTY) Then
            '    Call STOCKADJUST
            '    RSTSTKLESS!CLOSE_QTY = M_STOCK
            'End If
            
            i = i + 1
            grdSTOCKLESS.Rows = grdSTOCKLESS.Rows + 1
            grdSTOCKLESS.FixedRows = 1
            grdSTOCKLESS.TextMatrix(i, 0) = i
            grdSTOCKLESS.TextMatrix(i, 1) = RSTSTKLESS!ITEM_CODE
            grdSTOCKLESS.TextMatrix(i, 2) = RSTSTKLESS!ITEM_NAME
            grdSTOCKLESS.TextMatrix(i, 3) = RSTSTKLESS!CLOSE_QTY
            If IsNull(RSTSTKLESS!UNIT) Then
                grdSTOCKLESS.TextMatrix(i, 4) = 1
            Else
                grdSTOCKLESS.TextMatrix(i, 4) = RSTSTKLESS!UNIT
            End If
            If IsNull(RSTSTKLESS!REORDER_QTY) Then
                grdSTOCKLESS.TextMatrix(i, 5) = 1
            Else
                grdSTOCKLESS.TextMatrix(i, 5) = RSTSTKLESS!REORDER_QTY
            End If
        
SKIP:
        
        RSTORDER.Close
        Set RSTORDER = Nothing
        
        RSTSTKLESS.MoveNext
    Loop
    RSTSTKLESS.Close
    Set RSTSTKLESS = Nothing

    
    TMPFLAG = True
    grdSTOCKLESS_Click
    TXTUNIT.Visible = False
    Me.Left = 0
    Me.Top = 0
    Me.Height = 10000
    Me.Width = 10000
    Screen.MousePointer = vbNormal
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub grdSTOCKLESS_KeyPress(KeyAscii As Integer)
Dim i As Integer
    Select Case KeyAscii
        Case Asc("'")
            KeyAscii = 0
        Case vbKeyA To vbKeyZ, Asc("a") To Asc("z")
            i = 1
            Do Until i = grdSTOCKLESS.Rows
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If Asc(Mid(grdSTOCKLESS.TextMatrix(i, 2), 1, 1)) = KeyAscii Then Exit Do
                i = i + 1
            Loop
            If i = grdSTOCKLESS.Rows Then i = grdSTOCKLESS.Row
            grdSTOCKLESS.Row = i
            grdSTOCKLESS.TopRow = i
           
            grdSTOCKLESS.SetFocus
            grdSTOCKLESS_Click
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

Private Sub TXTUNIT_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTUNIT.Text) = 0 Then
                TXTUNIT.SetFocus
                Exit Sub
            End If
            CMDADDDIST.SetFocus
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

Private Sub CmbQty_KEyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            cmdadd.SetFocus
    End Select
End Sub

Private Function STOCKADJUST()
    Dim rststock As ADODB.Recordset
    Dim RSTITEMMAST As ADODB.Recordset
    Dim i As Integer
    
    i = 0
    If grdSTOCKLESS.Rows <= 1 Then Exit Function
    On Error GoTo eRRhAND
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT BAL_QTY from [RTRXFILE] where RTRXFILE.ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        i = i + rststock!BAL_QTY
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    M_STOCK = i
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & grdSTOCKLESS.TextMatrix(grdSTOCKLESS.Row, 1) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
    Exit Function
    
eRRhAND:
    MsgBox Err.Description
End Function

