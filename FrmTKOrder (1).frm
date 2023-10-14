VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmTKOrder 
   BackColor       =   &H00F8EDDE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Take Order from Customers"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15075
   Icon            =   "FrmTKOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   15075
   Begin VB.CommandButton Command3 
      Caption         =   "Truncate Table"
      Height          =   465
      Left            =   2745
      TabIndex        =   28
      Top             =   8205
      Width           =   1125
   End
   Begin VB.CommandButton CmdDeleteall 
      Caption         =   "Delete All"
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
      Left            =   9840
      TabIndex        =   14
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmddel 
      Caption         =   "&Delete Order"
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
      Left            =   11145
      TabIndex        =   15
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton CMDDETAILS 
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
      Height          =   435
      Left            =   15
      TabIndex        =   13
      Top             =   8190
      Width           =   1410
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
      Left            =   13710
      TabIndex        =   17
      Top             =   8190
      Width           =   1200
   End
   Begin VB.CommandButton cmddisplay 
      Caption         =   "&Display"
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
      Left            =   12435
      TabIndex        =   16
      Top             =   8190
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GrdCustomer 
      Height          =   5610
      Left            =   0
      TabIndex        =   11
      Top             =   2490
      Width           =   8340
      _ExtentX        =   14711
      _ExtentY        =   9895
      _Version        =   393216
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   350
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
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
   Begin MSFlexGridLib.MSFlexGrid GrdOrderlist 
      Height          =   5640
      Left            =   8370
      TabIndex        =   12
      Top             =   2475
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   9948
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   450
      BackColorFixed  =   0
      ForeColorFixed  =   65535
      BackColorBkg    =   12632256
      WordWrap        =   -1  'True
      AllowBigSelection=   0   'False
      FocusRect       =   2
      AllowUserResizing=   1
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00F8EDDE&
      Height          =   2565
      Left            =   0
      TabIndex        =   18
      Top             =   -75
      Width           =   9570
      Begin VB.TextBox txtBillNo 
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
         Left            =   4395
         TabIndex        =   29
         Top             =   2190
         Width           =   885
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8715
         TabIndex        =   9
         Top             =   1935
         Width           =   825
      End
      Begin VB.CommandButton cmddelete 
         Caption         =   "De&lete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7860
         TabIndex        =   8
         Top             =   1935
         Width           =   795
      End
      Begin VB.CommandButton cmdadd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   6990
         TabIndex        =   7
         Top             =   1935
         Width           =   810
      End
      Begin VB.TextBox txtQty 
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
         Left            =   6195
         MaxLength       =   6
         TabIndex        =   6
         Top             =   1935
         Width           =   765
      End
      Begin VB.TextBox TXTDEALER 
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
         Height          =   330
         Left            =   5820
         TabIndex        =   4
         Top             =   330
         Width           =   3735
      End
      Begin VB.TextBox TxtRemarks 
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
         Left            =   1035
         MaxLength       =   200
         TabIndex        =   3
         Top             =   1845
         Width           =   4245
      End
      Begin VB.TextBox TxtPhone 
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
         Left            =   1035
         MaxLength       =   35
         TabIndex        =   2
         Top             =   1500
         Width           =   2925
      End
      Begin VB.TextBox TxtBillName 
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
         Height          =   330
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   0
         Top             =   150
         Width           =   4215
      End
      Begin MSDataListLib.DataList DataList2 
         Height          =   1230
         Left            =   5820
         TabIndex        =   5
         Top             =   675
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   2170
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
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
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
         Left            =   45
         TabIndex        =   27
         Top             =   195
         Width           =   1110
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
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
         Left            =   5820
         TabIndex        =   24
         Top             =   1965
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
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
         Left            =   5805
         TabIndex        =   23
         Top             =   105
         Width           =   1245
      End
      Begin VB.Label lbldealer 
         Height          =   315
         Left            =   0
         TabIndex        =   22
         Top             =   765
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label flagchange 
         Height          =   315
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Index           =   35
         Left            =   45
         TabIndex        =   20
         Top             =   1845
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
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
         Index           =   47
         Left            =   30
         TabIndex        =   19
         Top             =   1515
         Width           =   1110
      End
      Begin MSForms.TextBox TxtBillAddress 
         Height          =   960
         Left            =   1035
         TabIndex        =   1
         Top             =   495
         Width           =   4215
         VariousPropertyBits=   -1400879077
         MaxLength       =   150
         BorderStyle     =   1
         Size            =   "7435;1693"
         SpecialEffect   =   0
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2475
      Left            =   9570
      TabIndex        =   25
      Top             =   -30
      Width           =   5445
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
         Left            =   3420
         TabIndex        =   26
         Top             =   600
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GrdOrder 
         Height          =   2400
         Left            =   0
         TabIndex        =   10
         Top             =   90
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   4233
         _Version        =   393216
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         BackColorFixed  =   0
         ForeColorFixed  =   65535
         BackColorBkg    =   12632256
         WordWrap        =   -1  'True
         AllowBigSelection=   0   'False
         FocusRect       =   2
         AllowUserResizing=   1
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
   End
End
Attribute VB_Name = "FrmTKOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub CMDADD_Click()
    Dim i As Integer
    If Val(TXTQTY.text) = 0 Then Exit Sub
    If Trim(TXTDEALER.text) = "" Then
        MsgBox "Please enter the Item Name", vbOKOnly, "Order"
        TXTDEALER.SetFocus
        Exit Sub
    End If
    For i = 1 To GrdOrder.rows - 1
        If Trim(GrdOrder.TextMatrix(i, 1)) = Trim(TXTDEALER.text) Then
            If MsgBox("This Item Already exists in Line No. " & i & " Do yo want to add this item", vbYesNo + vbDefaultButton2, "EzBiz") = vbNo Then
                TXTDEALER.SetFocus
                Exit Sub
            End If
            Exit For
        End If
    Next i
    GrdOrder.rows = GrdOrder.rows + 1
    GrdOrder.FixedRows = 1
    GrdOrder.TextMatrix(GrdOrder.rows - 1, 0) = GrdOrder.rows - 1
    GrdOrder.TextMatrix(GrdOrder.rows - 1, 1) = TXTDEALER.text
    GrdOrder.TextMatrix(GrdOrder.rows - 1, 2) = Val(TXTQTY.text)
    GrdOrder.TextMatrix(GrdOrder.rows - 1, 3) = DataList2.BoundText
    
    TXTDEALER.text = ""
    TXTQTY.text = ""
    TXTDEALER.SetFocus
End Sub

Private Sub cmddel_Click()
    If GrdCustomer.rows <= 1 Then Exit Sub
    If MsgBox("Are You Sure You want to delete the Order " & GrdCustomer.TextMatrix(GrdCustomer.Row, 1) & " from the list", vbYesNo + vbDefaultButton2, "Take Order....") = vbNo Then Exit Sub

    db.Execute "delete From ord_mast WHERE ord_no = " & GrdCustomer.TextMatrix(GrdCustomer.Row, 5) & ""
    db.Execute "delete From ord_trxfile WHERE ord_no = " & GrdCustomer.TextMatrix(GrdCustomer.Row, 5) & ""
    Call Fillgrid
    Call Fillgrid2
End Sub

Private Sub CmdDelete_Click()
    If GrdOrder.rows <= 1 Then Exit Sub
    If MsgBox("Are You Sure You want to delete item " & GrdOrder.TextMatrix(GrdOrder.Row, 1) & " from the list", vbYesNo + vbDefaultButton2, "Take Order....") = vbNo Then Exit Sub
    Dim i, selrow As Integer
    selrow = GrdOrder.Row
    For i = selrow To GrdOrder.rows - 2
        GrdOrder.TextMatrix(selrow, 0) = i
        GrdOrder.TextMatrix(selrow, 1) = GrdOrder.TextMatrix(i + 1, 1)
        GrdOrder.TextMatrix(selrow, 2) = GrdOrder.TextMatrix(i + 1, 2)
        selrow = selrow + 1
    Next i
    GrdOrder.rows = GrdOrder.rows - 1
'    For i = 1 To GrdOrder.Rows - 1
'        GrdOrder.TextMatrix(i, 0) = i
'    Next i
    
End Sub

Private Sub CmdDeleteAll_Click()
    If GrdCustomer.rows <= 1 Then Exit Sub
    If MsgBox("Are You Sure You want to delete all the Orders", vbYesNo + vbDefaultButton2, "Take Order....") = vbNo Then Exit Sub

    db.Execute "delete From ord_mast"
    db.Execute "delete From ord_trxfile"
    GrdCustomer.rows = 1
    GrdOrderlist.rows = 1
End Sub

Private Sub CMDDETAILS_Click()
    On Error GoTo ErrHand
    Dim i As Integer
    ReportNameVar = Rptpath & "RptTakeOrder"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
'    If OPTCUSTOMER.value = True Then
'        Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} ='DR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
'    Else
'        Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} ='DR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
'    End If
'
    Set CRXFormulaFields = Report.FormulaFields

    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox Err.Description
End Sub

Private Sub CmDDisplay_Click()
    Call Fillgrid
    Call Fillgrid2
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub CmdSave_Click()
    If GrdOrder.rows <= 1 Then Exit Sub
    If Trim(TxtBillName.text) = "" Then
        MsgBox "Please enter the name of the customer", vbOKCancel, "Order"
        TxtBillName.SetFocus
        Exit Sub
    End If
    If Trim(TxtPhone.text) = "" Then
        MsgBox "Please enter the phone number of the customer", vbOKCancel, "Order"
        TxtPhone.SetFocus
        Exit Sub
    End If
    
    Screen.MousePointer = vbHourglass
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTordtrxfile As ADODB.Recordset
    Dim i As Integer
        
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ord_mast WHERE ord_no= (SELECT MAX(ord_no) FROM ord_mast)", db, adOpenStatic, adLockOptimistic, adCmdText
    If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
        txtBillNo.text = 1
    Else
        txtBillNo.text = RSTTRXFILE!ord_no + 1
    End If
    db.BeginTrans
    RSTTRXFILE.AddNew
    RSTTRXFILE!ord_no = txtBillNo.text
    RSTTRXFILE!ACT_CODE = ""
    RSTTRXFILE!ACT_NAME = Trim(TxtBillName)
    RSTTRXFILE!act_address = Trim(TxtBillAddress.text)
    RSTTRXFILE!act_phone = Trim(TxtPhone.text)
'    RSTTRXFILE!ord_no = txtBillNo.Text
'    RSTTRXFILE!ord_no = txtBillNo.Text
'    RSTTRXFILE!C_USER_ID = frmLogin.rs!USER_ID
'    RSTTRXFILE!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTTRXFILE.Update
    
    For i = 1 To GrdOrder.rows - 1
        Set RSTordtrxfile = New ADODB.Recordset
        RSTordtrxfile.Open "Select * FROM ord_trxfile WHERE ord_no = " & Val(txtBillNo.text) & " AND line_no = " & Val(GrdOrder.TextMatrix(i, 0)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
        If (RSTordtrxfile.EOF And RSTordtrxfile.BOF) Then
            RSTordtrxfile.AddNew
            RSTordtrxfile!ord_no = txtBillNo.text
            RSTordtrxfile!line_no = Val(GrdOrder.TextMatrix(i, 0))
    '        RSTordtrxfile!C_USER_ID = frmLogin.rs!USER_ID
    '        RSTordtrxfile!CREATE_DATE = Format(Date, "DD/MM/YYYY")F
        End If
        RSTordtrxfile!ITEM_CODE = GrdOrder.TextMatrix(i, 3)
        RSTordtrxfile!ITEM_NAME = GrdOrder.TextMatrix(i, 1)
        RSTordtrxfile!ITEM_QTY = Val(GrdOrder.TextMatrix(i, 2))
        RSTordtrxfile.Update
    Next i
    db.CommitTrans
    
    RSTordtrxfile.Close
    Set RSTordtrxfile = Nothing
    
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    GrdOrder.rows = 1
    TXTDEALER.text = ""
    TXTQTY.text = ""
    TxtBillName.text = ""
    TxtBillAddress.text = ""
    TxtPhone.text = ""
    TXTREMARKS.text = ""
    Call Fillgrid
    Call Fillgrid2
    TxtBillName.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    If Err.Number <> -2147168237 Then
        MsgBox Err.Description
    End If
    On Error Resume Next
    db.RollbackTrans
End Sub

Private Sub Command3_Click()
'db.Execute "create table as_mast (PS_CODE VARCHAR(6) NOT NULL, AS_CODE VARCHAR(6) NOT NULL, PS_NAME TEXT(100), AS_NAME TEXT(150), AS_T_ADDRESS TEXT(255), AS_T_DISTRICT TEXT(30), AS_T_STATE TEXT(30), AS_T_PS TEXT(30), AS_T_CIRCLE TEXT(30), AS_T_SUB_DIVISION TEXT(30), AS_P_ADDRESS TEXT(255), AS_P_DISTRICT TEXT(30), AS_P_STATE TEXT(30), AS_P_PS TEXT(30), AS_P_CIRCLE TEXT(30), AS_P_SUB_DIVISION TEXT(30), AS_FATHER TEXT(100), AS_MARITAL_STATUS TEXT(1), AS_SPOUSE_NAME TEXT(30), AS_DOB DATE, AS_PHONE TEXT(15), AS_MOBILE TEXT(15), AS_CRIME_NO TEXT(5), AS_YEAR TEXT(4), AS_SEC_LAW TEXT(50), AS_PS TEXT(50), AS_CIRCLE TEXT(50), AS_SUB_DIVISION TEXT(50), AS_FLAG TEXT(1), AS_STATUS TEXT(1), C_USER_ID TEXT(4), C_USER_NAME TEXT(50), C_USER_DATE DATE, M_USER_ID TEXT(4), M_USER_NAME TEXT(50), M_USER_DATE DATE, PRIMARY KEY (PS_CODE, AS_CODE))"
    'On Error Resume Next
    db.Execute "DROP TABLE if exists `ord_mast`"
    db.Execute "FLUSH TABLES `ord_mast`"
    db.Execute "DROP TABLE if exists `ord_trxfile`"
    db.Execute "FLUSH TABLES `ord_trxfile`"
    
    db.Execute "DROP TABLE if exists `orders`"
    db.Execute "FLUSH TABLES `orders`"
    
    db.Execute "create table ord_mast (ord_no INT NOT NULL, act_code VARCHAR(25) NULL, act_name VARCHAR(100) NULL, act_address Text(200) NULL, act_phone TEXT(15), C_USER_ID TEXT(4), C_USER_NAME TEXT(50), C_USER_DATE DATE, M_USER_ID TEXT(4), M_USER_NAME TEXT(50), M_USER_DATE DATE, PRIMARY KEY (ord_no)) ENGINE = InnoDB"
    db.Execute "create table ord_trxfile (ord_no INT NOT NULL, line_no INT NOT NULL, item_code VARCHAR(25) NULL, item_name VARCHAR(200) NULL, item_qty int null, PRIMARY KEY (ord_no, line_no)) ENGINE = InnoDB"
    db.Execute "create table orders (ord_no INT NOT NULL, USER_ID VARCHAR(8) NULL, COMP_CODE VARCHAR(6) NOT NULL, act_code VARCHAR(25) NOT NULL, act_name VARCHAR(150) NULL, line_no INT NOT NULL, item_code VARCHAR(25) NULL, item_name VARCHAR(200) NULL, item_uprice double null, item_qty double null, item_tprice double null, c_date varchar(15) null, PRIMARY KEY (ord_no, COMP_CODE, ACT_CODE, line_no)) ENGINE = InnoDB"
    
    
'    db.Execute "CREATE  TABLE  tbltrxfile (TRX_TYPE varchar(2)  NOT  NULL , VCH_NO double NOT  NULL , TRX_YEAR varchar(4)  NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL,  "
'db.Execute "CREATE  TABLE  tbltrxfile (TRX_TYPE varchar(2)  NOT  NULL , VCH_NO double NOT  NULL , TRX_YEAR varchar(4)  NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL "
'db.Execute "CREATE  TABLE  tbltrxfile (TRX_TYPE varchar(2)  NOT  NULL , VCH_NO double NOT  NULL , TRX_YEAR varchar(4)  NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL "
'db.Execute "CREATE  TABLE  tbltrxfile (TRX_TYPE varchar(2)  NOT  NULL , VCH_NO double NOT  NULL , TRX_YEAR varchar(4)  NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL "
'db.Execute "CREATE  TABLE  tbltrxfile (TRX_TYPE varchar(2)  NOT  NULL , VCH_NO double NOT  NULL , TRX_YEAR varchar(4)  NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL "
    
    db.Execute "DROP TABLE if exists `tbletrxfile`"
    db.Execute "FLUSH TABLES `tbletrxfile`"
    db.Execute "CREATE  TABLE  tbletrxfile (table_code varchar(6) NOT  NULL , VCH_DATE date  DEFAULT NULL , LINE_NO double NOT  NULL , CATEGORY varchar(50)  DEFAULT NULL , ITEM_CODE varchar(20)  DEFAULT NULL , ITEM_NAME varchar(200)  DEFAULT NULL , QTY double  DEFAULT NULL , ITEM_COST double  DEFAULT NULL , MRP double  DEFAULT NULL , PTR double  DEFAULT NULL , P_RETAIL double  DEFAULT NULL , P_RETAILWOTAX double  DEFAULT NULL , SALES_PRICE double  DEFAULT NULL , SALES_TAX double  DEFAULT NULL , UNIT varchar(6)  DEFAULT NULL , VCH_DESC varchar(50)  DEFAULT NULL , REF_NO varchar(15)  DEFAULT NULL , ISSUE_QTY double  DEFAULT NULL , CST double  DEFAULT NULL , BAL_QTY double  DEFAULT NULL , TRX_TOTAL double  DEFAULT NULL , LINE_DISC double  DEFAULT NULL , " & _
    "SCHEME double  DEFAULT NULL , EXP_DATE text, FREE_QTY double  DEFAULT NULL , CREATE_DATE date  DEFAULT NULL , C_USER_ID varchar(8)  DEFAULT NULL , MODIFY_DATE date  DEFAULT NULL , M_USER_ID varchar(15)  DEFAULT NULL , CHECK_FLAG varchar(1)  DEFAULT NULL , AREA varchar(15)  DEFAULT NULL , MFGR varchar(50)  DEFAULT NULL , " & _
    "SALE_1_FLAG varchar(1)  DEFAULT NULL , COM_AMT double  DEFAULT NULL , COM_FLAG varchar(1)  DEFAULT NULL , LOOSE_FLAG varchar(1)  DEFAULT NULL , LOOSE_PACK int(11)  DEFAULT NULL , " & _
    "WARRANTY int(11)  DEFAULT NULL , WARRANTY_TYPE varchar(20)  DEFAULT NULL , PACK_TYPE varchar(20)  DEFAULT NULL , DN_NO double  DEFAULT NULL , DN_DATE datetime  DEFAULT NULL , " & _
    "RETAILER_PRICE double  DEFAULT NULL , PRINT_NAME varchar(200)  DEFAULT NULL , ST_RATE double  DEFAULT NULL , GROSS_AMOUNT double  DEFAULT NULL , DN_LINENO double  DEFAULT NULL , UN_BILL tinytext, CESS_AMT double  DEFAULT NULL , " & _
    "CESS_PER double  DEFAULT NULL , P_WS double  DEFAULT NULL , ITEM_SIZE varchar(5)  DEFAULT NULL , ITEM_COLOR varchar(10)  DEFAULT NULL , TAX_MODE tinytext, BARCODE tinytext, ITEM_SPEC varchar(300)  DEFAULT NULL ,  PRIMARY  KEY (table_code ,  LINE_NO)) ENGINE = InnoDB"
    
    
    db.Execute "DROP TABLE if exists `trnxtable`"
    db.Execute "FLUSH TABLES `trnxtable`"
    db.Execute "CREATE  TABLE  trnxtable (table_code varchar(6) NOT  NULL , SLSM_CODE varchar(1)  DEFAULT NULL , DISCOUNT double  DEFAULT NULL , ADD_AMOUNT double  DEFAULT NULL , FRIEGHT double  DEFAULT NULL , Handle double  DEFAULT NULL , AGENT_NAME varchar(50)  DEFAULT NULL , AGENT_CODE varchar(6)  DEFAULT NULL ,  PRIMARY  KEY (table_code)) ENGINE = InnoDB"

End Sub

Private Sub Form_Load()
    ACT_FLAG = True
    
    GrdOrder.TextMatrix(0, 0) = "SL"
    GrdOrder.TextMatrix(0, 1) = "Item Name"
    GrdOrder.TextMatrix(0, 2) = "Qty"
    
    GrdOrder.ColWidth(0) = 700
    GrdOrder.ColWidth(1) = 3600
    GrdOrder.ColWidth(2) = 900
    
    GrdOrder.ColAlignment(0) = 1
    GrdOrder.ColAlignment(1) = 1
    GrdOrder.ColAlignment(2) = 1
    Call Fillgrid
    Call Fillgrid2
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ACT_FLAG = False Then ACT_REC.Close
End Sub

Private Sub GrdCustomer_Click()
    Call Fillgrid2
End Sub

Private Sub TxtBillAddress_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            TxtPhone.SetFocus
        Case vbKeyEscape
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TxtBillName_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TxtBillName.text) = "" Then Exit Sub
            TxtBillAddress.SetFocus
        Case vbKeyEscape
        
    End Select
End Sub

Private Sub TXTDEALER_Change()
    On Error GoTo ErrHand
    If flagchange.Caption <> "1" Then
        If ACT_FLAG = True Then
            ACT_REC.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTDEALER.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        Else
            ACT_REC.Close
            ACT_REC.Open "Select * From ITEMMAST  WHERE ITEM_NAME Like '%" & Me.TXTDEALER.text & "%' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockReadOnly
            ACT_FLAG = False
        End If
        If (ACT_REC.EOF And ACT_REC.BOF) Then
            lbldealer.Caption = ""
        Else
            lbldealer.Caption = ACT_REC!ITEM_NAME
        End If
        Set Me.DataList2.RowSource = ACT_REC
        DataList2.ListField = "ITEM_NAME"
        DataList2.BoundColumn = "ITEM_CODE"
    End If
    Exit Sub
ErrHand:
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER_GotFocus()
    TXTDEALER.SelStart = 0
    TXTDEALER.SelLength = Len(TXTDEALER.text)
End Sub

Private Sub TXTDEALER_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TXTDEALER.text) = "" Then Exit Sub
            If DataList2.VisibleCount = 0 Then
                If MsgBox("Are You Sure You want to add an item not in the list", vbYesNo + vbDefaultButton2, "Take Order....") = vbNo Then Exit Sub
                TXTQTY.SetFocus
                Exit Sub
            End If
            DataList2.SetFocus
        Case vbKeyEscape
            TxtBillName.SetFocus
    End Select
End Sub

Private Sub TXTDEALER_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_Click()
    TXTDEALER.text = DataList2.text
    lbldealer.Caption = TXTDEALER.text
End Sub

Private Sub DataList2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If DataList2.text = "" Then Exit Sub
            If IsNull(DataList2.SelectedItem) Then
                MsgBox "Select Item From List", vbOKOnly, "Order"
                DataList2.SetFocus
                Exit Sub
            End If
            TXTQTY.SetFocus
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub DataList2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub DataList2_GotFocus()
    flagchange.Caption = 1
    TXTDEALER.text = lbldealer.Caption
    DataList2.text = TXTDEALER.text
    Call DataList2_Click
End Sub

Private Sub DataList2_LostFocus()
     flagchange.Caption = ""
End Sub

Private Sub TxtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            If Trim(TxtPhone.text) = "" Then Exit Sub
            TXTREMARKS.SetFocus
        Case vbKeyEscape
            TxtBillAddress.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(TXTQTY.text) = 0 Then Exit Sub
            CMDADD_Click
        Case vbKeyEscape
            TXTDEALER.SetFocus
    End Select
End Sub

Private Sub TXTQTY_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtremarks_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, 40
            TXTDEALER.SetFocus
        Case vbKeyEscape
            TxtPhone.SetFocus
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GrdOrder.Col
                Case 1  ' NAME
                    GrdOrder.TextMatrix(GrdOrder.Row, GrdOrder.Col) = Trim(TXTsample.text)
                    GrdOrder.Enabled = True
                    TXTsample.Visible = False
                    GrdOrder.SetFocus
                Case 2  ' QTY
                    GrdOrder.TextMatrix(GrdOrder.Row, GrdOrder.Col) = Val(TXTsample.text)
                    GrdOrder.Enabled = True
                    TXTsample.Visible = False
                    GrdOrder.SetFocus
            End Select
        Case vbKeyEscape
            TXTsample.Visible = False
            GrdOrder.SetFocus
    End Select
        Exit Sub
ErrHand:
    MsgBox Err.Description
    
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GrdOrder.Col
        Case 2
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 1
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
    End Select
End Sub

Private Sub GrdOrder_Click()
    If GrdOrder.Cols = 20 Then Exit Sub
    TXTsample.Visible = False
    GrdOrder.SetFocus
End Sub

Private Sub GrdOrder_KeyDown(KeyCode As Integer, Shift As Integer)
    If GrdOrder.rows = 1 Then Exit Sub
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GrdOrder.Col
                Case 1
                    TXTsample.MaxLength = 150
                    TXTsample.Visible = True
                    TXTsample.Top = GrdOrder.CellTop '+ 350
                    TXTsample.Left = GrdOrder.CellLeft '+ 50
                    TXTsample.Width = GrdOrder.CellWidth
                    TXTsample.Height = GrdOrder.CellHeight
                    TXTsample.text = Trim(GrdOrder.TextMatrix(GrdOrder.Row, GrdOrder.Col))
                    TXTsample.SetFocus
                Case 2
                    TXTsample.MaxLength = 6
                    TXTsample.Visible = True
                    TXTsample.Top = GrdOrder.CellTop '+ 350
                    TXTsample.Left = GrdOrder.CellLeft '+ 50
                    TXTsample.Width = GrdOrder.CellWidth
                    TXTsample.Height = GrdOrder.CellHeight
                    TXTsample.text = Val(GrdOrder.TextMatrix(GrdOrder.Row, GrdOrder.Col))
                    TXTsample.SetFocus
            End Select
            
    End Select
End Sub

Private Sub GrdOrder_Scroll()
    TXTsample.Visible = False
    GrdOrder.SetFocus
End Sub

Private Function Fillgrid()
    Dim rststock As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    
    i = 0
    Screen.MousePointer = vbHourglass
    
    GrdCustomer.rows = 1
    GrdCustomer.TextMatrix(i, 0) = "Sl"
    GrdCustomer.TextMatrix(i, 1) = "Cust Code"
    GrdCustomer.TextMatrix(i, 2) = "Customer"
    GrdCustomer.TextMatrix(i, 3) = "Address"
    GrdCustomer.TextMatrix(i, 4) = "Phone"
    GrdCustomer.TextMatrix(i, 5) = "Ord No"
    
    GrdCustomer.ColWidth(0) = 900
    GrdCustomer.ColWidth(1) = 0
    GrdCustomer.ColWidth(2) = 2000
    GrdCustomer.ColWidth(3) = 3500
    GrdCustomer.ColWidth(4) = 1600
    GrdCustomer.ColWidth(5) = 0
    
    GrdCustomer.ColAlignment(0) = 1
    GrdCustomer.ColAlignment(1) = 1
    GrdCustomer.ColAlignment(2) = 1
    GrdCustomer.ColAlignment(3) = 1
    GrdCustomer.ColAlignment(4) = 1
    GrdCustomer.ColAlignment(5) = 1
    
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM ord_mast ORDER BY ord_no", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + 1
        GrdCustomer.rows = GrdCustomer.rows + 1
        GrdCustomer.FixedRows = 1
        GrdCustomer.TextMatrix(i, 0) = i
        GrdCustomer.TextMatrix(i, 1) = IIf(IsNull(rststock!ACT_CODE), "", rststock!ACT_CODE)
        GrdCustomer.TextMatrix(i, 2) = IIf(IsNull(rststock!ACT_NAME), "", rststock!ACT_NAME)
        GrdCustomer.TextMatrix(i, 3) = IIf(IsNull(rststock!act_address), "", rststock!act_address)
        GrdCustomer.TextMatrix(i, 4) = IIf(IsNull(rststock!act_phone), "", rststock!act_phone)
        GrdCustomer.TextMatrix(i, 5) = IIf(IsNull(rststock!ord_no), "", rststock!ord_no)
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing

    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Function Fillgrid2()
    Dim rststock As ADODB.Recordset
    Dim i As Long

    On Error GoTo ErrHand
    
    i = 0
    
    GrdOrderlist.rows = 1
    GrdOrderlist.TextMatrix(i, 0) = "Sl"
    GrdOrderlist.TextMatrix(i, 1) = "Item Code"
    GrdOrderlist.TextMatrix(i, 2) = "Item Description"
    GrdOrderlist.TextMatrix(i, 3) = "Qty"
    GrdOrderlist.TextMatrix(i, 4) = "Line No"
    GrdOrderlist.TextMatrix(i, 5) = "Ord No"
    
    GrdOrderlist.ColWidth(0) = 900
    GrdOrderlist.ColWidth(1) = 0
    GrdOrderlist.ColWidth(2) = 2500
    GrdOrderlist.ColWidth(3) = 900
    GrdOrderlist.ColWidth(4) = 1000
    GrdOrderlist.ColWidth(5) = 1000
    
    GrdOrderlist.ColAlignment(0) = 1
    GrdOrderlist.ColAlignment(1) = 1
    GrdOrderlist.ColAlignment(2) = 1
    GrdOrderlist.ColAlignment(3) = 4
    GrdOrderlist.ColAlignment(4) = 1
    GrdOrderlist.ColAlignment(5) = 1
    
    If GrdCustomer.Row = 0 Then Exit Function
    Screen.MousePointer = vbHourglass
    Set rststock = New ADODB.Recordset
    rststock.Open "SELECT * FROM ord_trxfile where ord_no = " & GrdCustomer.TextMatrix(GrdCustomer.Row, 5) & " ORDER BY ord_no, line_no", db, adOpenForwardOnly
    Do Until rststock.EOF
        i = i + 1
        GrdOrderlist.rows = GrdOrderlist.rows + 1
        GrdOrderlist.FixedRows = 1
        GrdOrderlist.TextMatrix(i, 0) = i
        GrdOrderlist.TextMatrix(i, 1) = IIf(IsNull(rststock!ITEM_CODE), "", rststock!ITEM_CODE)
        GrdOrderlist.TextMatrix(i, 2) = IIf(IsNull(rststock!ITEM_NAME), "", rststock!ITEM_NAME)
        GrdOrderlist.TextMatrix(i, 3) = IIf(IsNull(rststock!ITEM_QTY), "", rststock!ITEM_QTY)
        GrdOrderlist.TextMatrix(i, 4) = IIf(IsNull(rststock!line_no), "", rststock!line_no)
        GrdOrderlist.TextMatrix(i, 5) = IIf(IsNull(rststock!ord_no), "", rststock!ord_no)
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

