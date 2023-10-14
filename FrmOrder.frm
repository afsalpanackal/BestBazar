VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmOrder2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order -II"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   19110
   Icon            =   "FrmOrder.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   19110
   Begin VB.PictureBox picChecked 
      Height          =   285
      Left            =   0
      Picture         =   "FrmOrder.frx":08CA
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.PictureBox picUnchecked 
      Height          =   285
      Left            =   330
      Picture         =   "FrmOrder.frx":0C0C
      ScaleHeight     =   225
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   30
      Visible         =   0   'False
      Width           =   285
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
      Height          =   450
      Left            =   30
      TabIndex        =   14
      Top             =   7965
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00F2E8DB&
      Height          =   1530
      Left            =   0
      TabIndex        =   8
      Top             =   -75
      Width           =   16545
      Begin VB.OptionButton OptCompany 
         BackColor       =   &H00F2E8DB&
         Caption         =   "Company"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   3840
         TabIndex        =   33
         Top             =   510
         Width           =   1275
      End
      Begin VB.OptionButton OptCategory 
         BackColor       =   &H00F2E8DB&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   3840
         TabIndex        =   32
         Top             =   135
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.TextBox TxtItem1 
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
         Left            =   14310
         TabIndex        =   5
         Top             =   135
         Width           =   2175
      End
      Begin VB.Frame Frameexclude 
         BackColor       =   &H00F2E8DB&
         Height          =   1020
         Left            =   12075
         TabIndex        =   23
         Top             =   480
         Width           =   4425
         Begin VB.OptionButton Option3 
            BackColor       =   &H00F2E8DB&
            Caption         =   "All Barcodes"
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
            Left            =   45
            TabIndex        =   26
            Top             =   690
            Width           =   4260
         End
         Begin VB.OptionButton Option2 
            BackColor       =   &H00F2E8DB&
            Caption         =   "Show same items with diff prices"
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
            Left            =   45
            TabIndex        =   25
            Top             =   390
            Value           =   -1  'True
            Width           =   4260
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00F2E8DB&
            Caption         =   "Show same items with diff barcodes"
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
            Left            =   45
            TabIndex        =   24
            Top             =   105
            Width           =   4335
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00F2E8DB&
         Height          =   1020
         Left            =   8535
         TabIndex        =   18
         Top             =   480
         Width           =   3510
         Begin VB.OptionButton OptAll 
            BackColor       =   &H00F2E8DB&
            Caption         =   "Display All"
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
            Left            =   60
            TabIndex        =   20
            Top             =   315
            Width           =   1335
         End
         Begin VB.OptionButton OptStock 
            BackColor       =   &H00F2E8DB&
            Caption         =   "Order Items Only"
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
            Left            =   1500
            TabIndex        =   19
            Top             =   300
            Value           =   -1  'True
            Width           =   1905
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
         Left            =   9660
         TabIndex        =   4
         Top             =   135
         Width           =   4620
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
         Left            =   5130
         TabIndex        =   2
         Top             =   135
         Width           =   3300
      End
      Begin VB.TextBox txtmonth2 
         Appearance      =   0  'Flat
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
         Height          =   390
         Left            =   1725
         TabIndex        =   0
         Top             =   180
         Width           =   1050
      End
      Begin VB.TextBox txtmonth1 
         Appearance      =   0  'Flat
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
         Height          =   390
         Left            =   1725
         TabIndex        =   1
         Top             =   585
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   9930
         TabIndex        =   9
         Top             =   1335
         Visible         =   0   'False
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarForeColor=   192
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   108855297
         CurrentDate     =   40498
      End
      Begin MSDataListLib.DataList DataList1 
         Height          =   780
         Left            =   5130
         TabIndex        =   3
         Top             =   480
         Width           =   3300
         _ExtentX        =   5821
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
         Index           =   3
         Left            =   8550
         TabIndex        =   17
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label LBLDEALER2 
         Height          =   315
         Left            =   0
         TabIndex        =   16
         Top             =   960
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Label FLAGCHANGE2 
         Height          =   315
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label3 
         BackColor       =   &H00F2E8DB&
         Caption         =   "Required for"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   180
         TabIndex        =   13
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label Label4 
         BackColor       =   &H00F2E8DB&
         Caption         =   "days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   2910
         TabIndex        =   12
         Top             =   660
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H00F2E8DB&
         Caption         =   "days"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Left            =   2895
         TabIndex        =   11
         Top             =   255
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00F2E8DB&
         Caption         =   "Sold for the last"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   300
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   1560
      End
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
      TabIndex        =   7
      Top             =   7950
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
      Left            =   12375
      TabIndex        =   6
      Top             =   7950
      Width           =   1200
   End
   Begin VB.Frame Frame3 
      Height          =   6570
      Left            =   0
      TabIndex        =   27
      Top             =   1380
      Width           =   19065
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
         Left            =   6945
         TabIndex        =   29
         Top             =   2730
         Visible         =   0   'False
         Width           =   1350
      End
      Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
         Height          =   6435
         Left            =   15
         TabIndex        =   28
         Top             =   105
         Width           =   19065
         _ExtentX        =   33629
         _ExtentY        =   11351
         _Version        =   393216
         Rows            =   1
         Cols            =   12
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   450
         BackColor       =   16777215
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
   End
   Begin VB.Label LBLTOTAL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount"
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
      Height          =   360
      Index           =   6
      Left            =   2640
      TabIndex        =   22
      Top             =   8040
      Width           =   2100
   End
   Begin VB.Label LBLTRXTOTAL 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   4425
      TabIndex        =   21
      Top             =   7980
      Width           =   2220
   End
End
Attribute VB_Name = "FrmOrder2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY_REC As New ADODB.Recordset
Dim PHY_FLAG As Boolean

Private Sub CmdDelete_Click()
    Dim slnos As Integer
    Dim RSTTEM As ADODB.Recordset
    If GRDSTOCK.Rows <= 1 Then Exit Sub
    If MsgBox("Are you sure to Delete?", vbYesNo, "Order.....") = vbNo Then
        GRDSTOCK.SetFocus
        Exit Sub
    End If
    On Error GoTo Errhand
    db.Execute "delete FROM TEMPSTK"
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK", db, adOpenStatic, adLockOptimistic, adCmdText
    For slnos = 1 To GRDSTOCK.Rows - 1
        RSTTEM.AddNew
        RSTTEM!OPQTY = Val(GRDSTOCK.TextMatrix(slnos, 1))
        RSTTEM!ITEM_CODE = GRDSTOCK.TextMatrix(slnos, 2)
        RSTTEM!ITEM_NAME = Trim(Mid((GRDSTOCK.TextMatrix(slnos, 3)), 1, 200))
        RSTTEM!INQTY = Val(GRDSTOCK.TextMatrix(slnos, 4))
        RSTTEM!OUTQTY = Val(GRDSTOCK.TextMatrix(slnos, 5))
        RSTTEM!CLOSE_QTY = Val(GRDSTOCK.TextMatrix(slnos, 6))
        RSTTEM!CLOSE_VAL = Val(GRDSTOCK.TextMatrix(slnos, 7))
        RSTTEM!DIFF_QTY = Trim(Mid((GRDSTOCK.TextMatrix(slnos, 8)), 1, 100))
        RSTTEM!ITEM_COST = Val(GRDSTOCK.TextMatrix(slnos, 9))
        RSTTEM!BARCODE = Trim(GRDSTOCK.TextMatrix(slnos, 10))
        If GRDSTOCK.TextMatrix(slnos, 11) = "Y" Then 'GoTo SKIP
            RSTTEM!CHECK_FLAG = "N"
        Else
            RSTTEM!CHECK_FLAG = "Y"
        End If
        RSTTEM.Update
SKIP:
    Next slnos
    RSTTEM.Close
    Set RSTTEM = Nothing
    
    GRDSTOCK.Rows = 1
    slnos = 1
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK WHERE CHECK_FLAG ='Y' ORDER BY OPQTY", db, adOpenStatic, adLockReadOnly
    Do Until RSTTEM.EOF
        With GRDSTOCK
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            .TextMatrix(slnos, 1) = slnos
            .TextMatrix(slnos, 2) = RSTTEM!ITEM_CODE
            .TextMatrix(slnos, 3) = RSTTEM!ITEM_NAME
            .TextMatrix(slnos, 4) = RSTTEM!INQTY
            .TextMatrix(slnos, 5) = RSTTEM!OUTQTY
            .TextMatrix(slnos, 6) = RSTTEM!CLOSE_QTY
            .TextMatrix(slnos, 7) = RSTTEM!CLOSE_VAL
            .TextMatrix(slnos, 8) = RSTTEM!DIFF_QTY
            .TextMatrix(slnos, 9) = RSTTEM!ITEM_COST
            .TextMatrix(slnos, 10) = RSTTEM!BARCODE
            .Row = slnos: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
            Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            .TextMatrix(slnos, 11) = "N"
        End With
        'RSTTEM!CHECK_FLAG = "N"
        slnos = slnos + 1
        RSTTEM.MoveNext
    Loop
    RSTTEM.Close
    Set RSTTEM = Nothing
    GRDSTOCK.SetFocus
    
    LBLTRXTOTAL.Caption = ""
    Dim i As Long
    For i = 1 To GRDSTOCK.Rows - 1
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
    Next i
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDDETAILS_Click()
    Dim RSTTEM As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
    db.Execute "delete  FROM TEMPSTK"
    
    
    GRDSTOCK.TextMatrix(0, 1) = "SL"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 3) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 4) = "SOLD"
    GRDSTOCK.TextMatrix(0, 5) = "SHELF"
    GRDSTOCK.TextMatrix(0, 6) = "RQD QTY"
    GRDSTOCK.TextMatrix(0, 7) = "MRP"
    GRDSTOCK.TextMatrix(0, 8) = "SUPPLIER"
    
    If GRDSTOCK.Rows <= 1 Then Exit Sub
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To GRDSTOCK.Rows - 1
        RSTTEM.AddNew
        RSTTEM!ITEM_CODE = GRDSTOCK.TextMatrix(i, 2)
        RSTTEM!ITEM_NAME = Trim(Mid((GRDSTOCK.TextMatrix(i, 3)), 1, 200))
        RSTTEM!INQTY = Val(GRDSTOCK.TextMatrix(i, 4))
        RSTTEM!OUTQTY = Val(GRDSTOCK.TextMatrix(i, 5))
        RSTTEM!CLOSE_QTY = Val(GRDSTOCK.TextMatrix(i, 6))
        RSTTEM!CLOSE_VAL = Val(GRDSTOCK.TextMatrix(i, 7))
        RSTTEM!DIFF_QTY = Trim(Mid((GRDSTOCK.TextMatrix(i, 8)), 1, 100))
        RSTTEM!ITEM_COST = Val(GRDSTOCK.TextMatrix(i, 9))
        RSTTEM!BARCODE = Val(GRDSTOCK.TextMatrix(i, 10))
        RSTTEM!OPQTY = Val(GRDSTOCK.TextMatrix(i, 1))
        'RSTTEM!CHECK_FLAG = "N"
        'RSTTEM!OPVAL = 0 'GRDSTOCK.TextMatrix(i, 4)
        'RSTTEM!INQTY_VAL = GRDSTOCK.TextMatrix(i, 4)
'        RSTTEM!OUTQTY_VAL = GRDSTOCK.TextMatrix(i, 6)
'        RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 7)
'        RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 8)
        
        'RSTTEM!DIFF_QTY = 0 'GRDSTOCK.TextMatrix(i, 11)
        'RSTTEM!DIFF_VAL = 0 'GRDSTOCK.TextMatrix(i, 11)
        RSTTEM.Update
    Next i
    RSTTEM.Close
    Set RSTTEM = Nothing
    
    frmreport.Caption = "STOCK RE-ORDER"
    ReportNameVar = Rptpath & "RptOrder"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
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
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'" & MDIMAIN.StatusBar.Panels(5).Text & "'"
    Next

    Call GENERATEREPORT
    
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_Click()
    If OptAll.value = True Then
        Call Order
    Else
        If Option1.value = True Then
            Call Order4
        ElseIf Option2.value = True Then
            Call Order_1
        Else
            Call Order
        End If
    End If
    db.Execute "delete  FROM TEMPSTK"
End Sub

Private Sub CMDEXIT_Click()
    Unload Me
End Sub

Private Sub CMDUNDO_Click()
     Dim slnos As Integer
    Dim RSTTEM As ADODB.Recordset
    'If GRDSTOCK.Rows <= 1 Then Exit Sub
    If MsgBox("Are you sure to Restore?", vbYesNo, "Order.....") = vbNo Then
        GRDSTOCK.SetFocus
        Exit Sub
    End If
    On Error GoTo Errhand
    'db.Execute "delete  FROM TEMPSTK"
    'GRDSTOCK.Rows = 1
    slnos = 1
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK ORDER BY OPQTY", db, adOpenStatic, adLockReadOnly
    If RSTTEM.RecordCount = 0 Then
        GoTo SKIP
    Else
        GRDSTOCK.Rows = 1
    End If
    Do Until RSTTEM.EOF
        With GRDSTOCK
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            .TextMatrix(slnos, 1) = slnos
            .TextMatrix(slnos, 2) = RSTTEM!ITEM_CODE
            .TextMatrix(slnos, 3) = RSTTEM!ITEM_NAME
            .TextMatrix(slnos, 4) = RSTTEM!INQTY
            .TextMatrix(slnos, 5) = RSTTEM!OUTQTY
            .TextMatrix(slnos, 6) = RSTTEM!CLOSE_QTY
            .TextMatrix(slnos, 7) = RSTTEM!CLOSE_VAL
            .TextMatrix(slnos, 8) = RSTTEM!DIFF_QTY
            .TextMatrix(slnos, 9) = RSTTEM!ITEM_COST
            .TextMatrix(slnos, 10) = RSTTEM!BARCODE
            .Row = slnos: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
            Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
            .TextMatrix(slnos, 11) = "N"
        End With
        'RSTTEM!CHECK_FLAG = "N"
        slnos = slnos + 1
        RSTTEM.MoveNext
    Loop
SKIP:
    RSTTEM.Close
    Set RSTTEM = Nothing
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub Form_Load()
    
    GRDSTOCK.TextMatrix(0, 0) = ""
    GRDSTOCK.TextMatrix(0, 1) = "SL"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 3) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 4) = "SOLD"
    GRDSTOCK.TextMatrix(0, 5) = "SHELF"
    GRDSTOCK.TextMatrix(0, 6) = "RQD QTY"
    GRDSTOCK.TextMatrix(0, 7) = "MRP"
    GRDSTOCK.TextMatrix(0, 8) = "SUPPLIER"
    GRDSTOCK.TextMatrix(0, 9) = "COST"
    GRDSTOCK.TextMatrix(0, 10) = "BARCODE"
    GRDSTOCK.TextMatrix(0, 11) = ""
    
    GRDSTOCK.ColWidth(0) = 300
    GRDSTOCK.ColWidth(1) = 700
    GRDSTOCK.ColWidth(2) = 1000
    GRDSTOCK.ColWidth(3) = 7000
    GRDSTOCK.ColWidth(4) = 900
    GRDSTOCK.ColWidth(5) = 900
    GRDSTOCK.ColWidth(6) = 1200
    GRDSTOCK.ColWidth(7) = 1000
    GRDSTOCK.ColWidth(8) = 2900
    GRDSTOCK.ColWidth(9) = 1000
    GRDSTOCK.ColWidth(10) = 2700
    GRDSTOCK.ColWidth(11) = 0
        
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 1
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    GRDSTOCK.ColAlignment(8) = 1
    GRDSTOCK.ColAlignment(9) = 4
    GRDSTOCK.ColAlignment(10) = 1
    PHY_FLAG = True
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    Left = 0
    Top = 0
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHY_FLAG = False Then PHY_REC.Close
   'Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_Click()
    TXTsample.Visible = False
    Dim oldx, oldy, cell2text As String, strTextCheck As String
    If GRDSTOCK.Rows = 1 Then Exit Sub
    With GRDSTOCK
        oldx = .Col
        oldy = .Row
        .Row = oldy: .Col = 0: .CellPictureAlignment = 4
            'If GRDSTOCK.Col = 0 Then
                If GRDSTOCK.CellPicture = picChecked Then
                    Set GRDSTOCK.CellPicture = picUnchecked
                    '.Col = .Col + 2  ' I use data that is in column #1, usually an Index or ID #
                    'strTextCheck = .Text
                    ' When you de-select a CheckBox, we need to strip out the #
                    'strChecked = strChecked & strTextCheck & ","
                    ' Don't forget to strip off the trailing , before passing the string
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 11) = "Y"
                Else
                    Set GRDSTOCK.CellPicture = picChecked
                    '.Col = .Col + 2
                    'strTextCheck = .Text
                    'strChecked = Replace(strChecked, strTextCheck & ",", "")
                    'Debug.Print strChecked
                    .TextMatrix(.Row, 11) = "N"
                End If
            'End If
        .Col = oldx
        .Row = oldy
    End With
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 32
            Call GRDSTOCK_Click
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 4, 5, 6, 7, 8
                    TXTsample.Visible = True
                    TXTsample.Top = GRDSTOCK.CellTop + 90
                    TXTsample.Left = GRDSTOCK.CellLeft ' + 50
                    TXTsample.Width = GRDSTOCK.CellWidth
                    TXTsample.Height = GRDSTOCK.CellHeight
                    TXTsample.Text = GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col)
                    TXTsample.SetFocus
            End Select
        Case 68, 46
            Call CmdDelete_Click
            'On Error Resume Next
'            If MsgBox("Are you sure to Delete?", vbYesNo, "Order.....") = vbNo Then
'                GRDSTOCK.SetFocus
'                Exit Sub
'            End If
'            Dim del_count, count As Integer
'            count = 0
'            For del_count = 1 To GRDSTOCK.Rows - 1
'                If GRDSTOCK.TextMatrix(del_count, 11) = "N" Then GoTo SKIP
''                If del_count <> GRDSTOCK.Rows - 1 Then
''                    If Mid(GRDSTOCK.TextMatrix(del_count + 1, 3), 1, 3) = "==>" And Mid(GRDSTOCK.TextMatrix(del_count, 3), 1, 3) <> "==>" Then
''                        MsgBox "Cannot Delete since sub entries exists", vbOKOnly, "Order"
''                        GRDSTOCK.SetFocus
''                        Exit Sub
''                    End If
''                End If
'                If Mid(GRDSTOCK.TextMatrix(del_count, 3), 1, 3) <> "==>" Then LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) - (Val(GRDSTOCK.TextMatrix(del_count, 9)) * Val(GRDSTOCK.TextMatrix(del_count, 6)))
'                count = count + 1
'                For i = del_count To GRDSTOCK.Rows - 2
'                    GRDSTOCK.TextMatrix(i, 0) = "" 'GRDSTOCK.TextMatrix(i + 1, 1)
'                    GRDSTOCK.TextMatrix(i, 1) = i '- 1
'                    GRDSTOCK.TextMatrix(i, 2) = GRDSTOCK.TextMatrix(i + 1, 2)
'                    GRDSTOCK.TextMatrix(i, 3) = GRDSTOCK.TextMatrix(i + 1, 3)
'                    GRDSTOCK.TextMatrix(i, 4) = GRDSTOCK.TextMatrix(i + 1, 4)
'                    GRDSTOCK.TextMatrix(i, 6) = GRDSTOCK.TextMatrix(i + 1, 6)
'                    GRDSTOCK.TextMatrix(i, 5) = GRDSTOCK.TextMatrix(i + 1, 5)
'                    GRDSTOCK.TextMatrix(i, 7) = GRDSTOCK.TextMatrix(i + 1, 7)
'                    GRDSTOCK.TextMatrix(i, 8) = GRDSTOCK.TextMatrix(i + 1, 8)
'                    GRDSTOCK.TextMatrix(i, 9) = GRDSTOCK.TextMatrix(i + 1, 9)
'                    GRDSTOCK.TextMatrix(i, 10) = GRDSTOCK.TextMatrix(i + 1, 10)
'                    GRDSTOCK.TextMatrix(i, 11) = GRDSTOCK.TextMatrix(i + 1, 11)
'                Next i
'                GRDSTOCK.Rows = GRDSTOCK.Rows - count
'                If GRDSTOCK.Rows <= 1 Then LBLTRXTOTAL.Caption = ""
'                If GRDSTOCK.TextMatrix(del_count, 11) = "Y" Then
'                    del_count = del_count - 1
'                    'count = count - 1
'                End If
'SKIP:
'            Next del_count
'
'            GRDSTOCK.SetFocus
        Case 114
            sitem = UCase(InputBox("Item Name..?", "ZERO STOCK"))
            For i = 1 To GRDSTOCK.Rows - 1
                    If Mid(GRDSTOCK.TextMatrix(i, 3), 1, Len(sitem)) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub GRDSTOCK_Scroll()
    TXTsample.Visible = False
End Sub

Private Sub OptAll_Click()
    If OptAll.value = True Then
        Frameexclude.Visible = False
    Else
        Frameexclude.Visible = True
    End If
End Sub

Private Sub optCategory_Click()
    Call TXTDEALER2_Change
    TXTDEALER2.SetFocus
End Sub

Private Sub OptCompany_Click()
    Call TXTDEALER2_Change
    TXTDEALER2.SetFocus
End Sub

Private Sub OptStock_Click()
    If OptAll.value = True Then
        Frameexclude.Visible = False
    Else
        Frameexclude.Visible = True
    End If
End Sub

Private Sub TXTDEALER2_Change()
    
    
    On Error GoTo Errhand
    If FLAGCHANGE2.Caption <> "1" Then
        If optCategory.value = True Then
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
Errhand:
    MsgBox Err.Description
End Sub

Private Sub TXTDEALER2_GotFocus()
    TXTDEALER2.SelStart = 0
    TXTDEALER2.SelLength = Len(TXTDEALER2.Text)
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
            CmdDisplay.SetFocus
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
    FLAGCHANGE2.Caption = 1
    TXTDEALER2.Text = LBLDEALER2.Caption
    DataList1.Text = TXTDEALER2.Text
    Call DataList1_Click
End Sub

Private Sub DataList1_LostFocus()
     FLAGCHANGE2.Caption = ""
End Sub

Private Sub tXTMEDICINE_GotFocus()
    tXTMEDICINE.SelStart = 0
    tXTMEDICINE.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub tXTMEDICINE_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            TxtItem1.SetFocus
    End Select

End Sub

Private Function Order()
    Dim rststock, RSTRTRXFILE, RSTRTRXFILE2, RSTRTRXFILE3, RSTSUPPLIER, RSTSUPPLIER2 As ADODB.Recordset
    Dim i As Long
    Dim month1_stock, month2_stock, close_stock, RQD_QTY As Double
    'PHY_FLAG = True
    
    If Val(txtmonth2.Text) = 0 Or Val(txtmonth1.Text) = 0 Then
        MsgBox "Please enter proper values", vbOKOnly, "Order"
        txtmonth2.SetFocus
        Exit Function
    End If
    LBLTRXTOTAL.Caption = ""
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    'Screen.MousePointer = vbHourglass
    'GRDSTOCK.TextMatrix(0, 4) = "Sold last " & Val(txtmonth2.Text) & " days"
    'GRDSTOCK.TextMatrix(0, 6) = "Required for " & Val(txtmonth1.Text) & " days"
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    If DataList1.BoundText = "" Then
        rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Else
        If optCategory.value = True Then
            rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL, CATEGORY FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL, MANUFACTURER FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        month1_stock = 0
        month2_stock = 0
        close_stock = 0
        RQD_QTY = 0
        
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE.EOF
            close_stock = close_stock + RSTRTRXFILE!BAL_QTY
            RSTRTRXFILE.MoveNext
        Loop
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE.EOF
            month1_stock = month1_stock + RSTRTRXFILE!QTY
            RSTRTRXFILE.MoveNext
        Loop
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
        Set RSTRTRXFILE2 = New ADODB.Recordset
        RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE2.EOF
            month2_stock = month2_stock + RSTRTRXFILE2!QTY
            RSTRTRXFILE2.MoveNext
        Loop
        RSTRTRXFILE2.Close
        
        If close_stock < 0 Then
            RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
        Else
            RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
        End If
        If RQD_QTY < 0 Then RQD_QTY = "0"
        If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
        If OptStock.value = True And RQD_QTY <= 0 Then GoTo SKIP
        
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 1) = i
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 3) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 4) = month2_stock
        GRDSTOCK.TextMatrix(i, 5) = close_stock
        If Val(GRDSTOCK.TextMatrix(i, 5)) < 0 Then
            GRDSTOCK.TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
        Else
            GRDSTOCK.TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
        End If
        GRDSTOCK.TextMatrix(i, 7) = IIf(IsNull(rststock!P_RETAIL), "", Format(rststock!P_RETAIL, "0.00"))
        If Val(GRDSTOCK.TextMatrix(i, 6)) < 0 Then GRDSTOCK.TextMatrix(i, 6) = "0"
        If Val(GRDSTOCK.TextMatrix(i, 5)) <= 0 And Val(GRDSTOCK.TextMatrix(i, 6)) = 0 Then GRDSTOCK.TextMatrix(i, 6) = "1"
        GRDSTOCK.TextMatrix(i, 8) = "Opening Stock"
        Set RSTSUPPLIER = New ADODB.Recordset
        RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
        If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
            GRDSTOCK.TextMatrix(i, 8) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
             GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
        Else
            Set RSTSUPPLIER2 = New ADODB.Recordset
            RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                GRDSTOCK.TextMatrix(i, 8) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
            Else
                GRDSTOCK.TextMatrix(i, 9) = ""
            End If
            RSTSUPPLIER2.Close
            Set RSTSUPPLIER2 = Nothing
        End If
        RSTSUPPLIER.Close
        Set RSTSUPPLIER = Nothing
        'GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
SKIP:
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLTRXTOTAL.Caption = ""
    For i = 1 To GRDSTOCK.Rows - 1
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
    Next i
    
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Function

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Function Order_1()
    Dim rststock, RSTRTRXFILE, RSTRTRXFILE2, RSTRTRXFILE3, RSTRTRXFILE4, RSTRTRXFILE5, RSTSUPPLIER, RSTSUPPLIER2 As ADODB.Recordset
    Dim i, COUNT As Integer
    Dim month1_stock, month2_stock, close_stock, RQD_QTY As Double
    'PHY_FLAG = True
    
    If Val(txtmonth2.Text) = 0 Or Val(txtmonth1.Text) = 0 Then
        MsgBox "Please enter proper values", vbOKOnly, "Order"
        txtmonth2.SetFocus
        Exit Function
    End If
    
    LBLTRXTOTAL.Caption = ""
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    'Screen.MousePointer = vbHourglass
    'GRDSTOCK.TextMatrix(0, 3) = "Sold last " & Val(txtmonth2.Text) & " days"
    'GRDSTOCK.TextMatrix(0, 5) = "Required for " & Val(txtmonth1.Text) & " days"
    
    Dim Item_found As Boolean
    Dim STOCKQTY As Double
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    If DataList1.BoundText = "" Then
        rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Else
        If optCategory.value = True Then
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        
        COUNT = 0
        Set RSTRTRXFILE4 = New ADODB.Recordset
        RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, P_RETAIL FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE4.EOF
            If RSTRTRXFILE4.RecordCount = 1 Then
                COUNT = 0
                Exit Do
            End If
            month1_stock = 0
            month2_stock = 0
            close_stock = 0
            RQD_QTY = 0
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                month1_stock = month1_stock + RSTRTRXFILE!QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
            Set RSTRTRXFILE2 = New ADODB.Recordset
            RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE2.EOF
                month2_stock = month2_stock + RSTRTRXFILE2!QTY
                RSTRTRXFILE2.MoveNext
            Loop
            RSTRTRXFILE2.Close
            
            If close_stock < 0 Then
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
            Else
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
            End If
            If RQD_QTY < 0 Then RQD_QTY = "0"
            If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            If RQD_QTY > 0 Then COUNT = COUNT + 1
            RSTRTRXFILE4.MoveNext
        Loop
        RSTRTRXFILE4.Close
        Set RSTRTRXFILE4 = Nothing
            
        Item_found = False
        Set RSTRTRXFILE4 = New ADODB.Recordset
        RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, P_RETAIL FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE4.EOF
            month1_stock = 0
            month2_stock = 0
            close_stock = 0
            RQD_QTY = 0
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND VCH_DATE >= '" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                month1_stock = month1_stock + RSTRTRXFILE!QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
            Set RSTRTRXFILE2 = New ADODB.Recordset
            RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE2.EOF
                month2_stock = month2_stock + RSTRTRXFILE2!QTY
                RSTRTRXFILE2.MoveNext
            Loop
            RSTRTRXFILE2.Close
            
            If close_stock < 0 Then
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
            Else
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
            End If
            If RQD_QTY < 0 Then RQD_QTY = "0"
            If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            If OptStock.value = True And RQD_QTY <= 0 Then
                If COUNT = 0 Then GoTo SKIP2
            End If
            i = i + 1
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            Item_found = True
            GRDSTOCK.FixedRows = 1
            With GRDSTOCK
                .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                .TextMatrix(i, 11) = "N"
                .TextMatrix(i, 1) = i
                .TextMatrix(i, 2) = rststock!ITEM_CODE
                .TextMatrix(i, 3) = rststock!ITEM_NAME
                '.TextMatrix(i, 3) = "==>" & rststock!ITEM_NAME
                .TextMatrix(i, 4) = month2_stock
                .TextMatrix(i, 5) = close_stock
                If Val(.TextMatrix(i, 5)) < 0 Then
                    .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                Else
                    .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                End If
                'If COUNT = 0 Then
                .TextMatrix(i, 7) = IIf(IsNull(RSTRTRXFILE4!P_RETAIL), "", Format(RSTRTRXFILE4!P_RETAIL, "0.00"))
                If Val(.TextMatrix(i, 6)) < 0 Then .TextMatrix(i, 6) = "0"
                If Val(.TextMatrix(i, 5)) <= 0 And Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = "1"
                .TextMatrix(i, 8) = "Opening Stock"
                Set RSTSUPPLIER = New ADODB.Recordset
                RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                    .TextMatrix(i, 8) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                    .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
                Else
                    Set RSTSUPPLIER2 = New ADODB.Recordset
                    RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                        .TextMatrix(i, 8) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                        .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
                    Else
                        .TextMatrix(i, 9) = ""
                    End If
                    RSTSUPPLIER2.Close
                    Set RSTSUPPLIER2 = Nothing
                End If
                RSTSUPPLIER.Close
                Set RSTSUPPLIER = Nothing
                LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(.TextMatrix(i, 9)) * Val(.TextMatrix(i, 6)))
            End With
            Item_found = True
            Exit Do
SKIP2:
            RSTRTRXFILE4.MoveNext
        Loop
        RSTRTRXFILE4.Close
        Set RSTRTRXFILE4 = Nothing
        
        If COUNT > 0 Then
            Set RSTRTRXFILE4 = New ADODB.Recordset
            RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, P_RETAIL FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE4.EOF
                If RSTRTRXFILE4.RecordCount = 1 Then Exit Do
                month1_stock = 0
                month2_stock = 0
                close_stock = 0
                RQD_QTY = 0
                
                Set RSTRTRXFILE = New ADODB.Recordset
                RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE.EOF
                    close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                    RSTRTRXFILE.MoveNext
                Loop
                RSTRTRXFILE.Close
                Set RSTRTRXFILE = Nothing
                
                DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
                Set RSTRTRXFILE = New ADODB.Recordset
                RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE.EOF
                    month1_stock = month1_stock + RSTRTRXFILE!QTY
                    RSTRTRXFILE.MoveNext
                Loop
                RSTRTRXFILE.Close
                Set RSTRTRXFILE = Nothing
                
                DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
                Set RSTRTRXFILE2 = New ADODB.Recordset
                RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE2.EOF
                    month2_stock = month2_stock + RSTRTRXFILE2!QTY
                    RSTRTRXFILE2.MoveNext
                Loop
                RSTRTRXFILE2.Close
                
                If close_stock < 0 Then
                    RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                Else
                    RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                End If
                If RQD_QTY < 0 Then RQD_QTY = "0"
                If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            
                i = i + 1
                GRDSTOCK.Rows = GRDSTOCK.Rows + 1
                Item_found = True
                GRDSTOCK.FixedRows = 1
                With GRDSTOCK
                    .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
                    Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                    .TextMatrix(i, 11) = "N"
                    .TextMatrix(i, 1) = i
                    .TextMatrix(i, 2) = rststock!ITEM_CODE
                    .TextMatrix(i, 3) = "==>" & rststock!ITEM_NAME
                    .TextMatrix(i, 4) = month2_stock
                    .TextMatrix(i, 5) = close_stock
                    If Val(.TextMatrix(i, 5)) < 0 Then
                        .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                    Else
                        .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                    End If
                    .TextMatrix(i, 7) = IIf(IsNull(RSTRTRXFILE4!P_RETAIL), "", Format(RSTRTRXFILE4!P_RETAIL, "0.00"))
                    If Val(.TextMatrix(i, 6)) < 0 Then .TextMatrix(i, 6) = "0"
                    If Val(.TextMatrix(i, 5)) <= 0 And Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = "1"
                    .TextMatrix(i, 8) = "Opening Stock"
                    Set RSTSUPPLIER = New ADODB.Recordset
                    RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                        .TextMatrix(i, 8) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                        .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
                        .TextMatrix(i, 10) = IIf(IsNull(RSTSUPPLIER!BARCODE), "", RSTSUPPLIER!BARCODE)
                    Else
                        Set RSTSUPPLIER2 = New ADODB.Recordset
                        RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND P_RETAIL = " & RSTRTRXFILE4!P_RETAIL & " ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                        If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                            .TextMatrix(i, 8) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                            .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
                            .TextMatrix(i, 10) = IIf(IsNull(RSTSUPPLIER2!BARCODE), "", RSTSUPPLIER2!BARCODE)
                        Else
                            .TextMatrix(i, 9) = ""
                        End If
                        RSTSUPPLIER2.Close
                        Set RSTSUPPLIER2 = Nothing
                    End If
                    RSTSUPPLIER.Close
                    Set RSTSUPPLIER = Nothing
                End With
                RSTRTRXFILE4.MoveNext
            Loop
            RSTRTRXFILE4.Close
            Set RSTRTRXFILE4 = Nothing
        End If
        'GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLTRXTOTAL.Caption = ""
    For i = 1 To GRDSTOCK.Rows - 1
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
    Next i
    
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Function

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function


Private Function Order_2()
    Dim rststock, RSTRTRXFILE, RSTRTRXFILE2, RSTRTRXFILE3, RSTSUPPLIER, RSTSUPPLIER2 As ADODB.Recordset
    Dim i As Long
    Dim month1_stock, month2_stock, close_stock, RQD_QTY As Double
    'PHY_FLAG = True
    
    If Val(txtmonth2.Text) = 0 Or Val(txtmonth1.Text) = 0 Then
        MsgBox "Please enter proper values", vbOKOnly, "Order"
        txtmonth2.SetFocus
        Exit Function
    End If
    LBLTRXTOTAL.Caption = ""
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    'Screen.MousePointer = vbHourglass
    'GRDSTOCK.TextMatrix(0, 3) = "Sold last " & Val(txtmonth2.Text) & " days"
    'GRDSTOCK.TextMatrix(0, 5) = "Required for " & Val(txtmonth1.Text) & " days"
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    If DataList1.BoundText = "" Then
        rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL, BARCODE FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Else
        If optCategory.value = True Then
            rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL, BARCODE, CATEGORY FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            rststock.Open "SELECT DISTINCT ITEM_CODE, ITEM_NAME, P_RETAIL, BARCODE, MANUFACTURER FROM RTRXFILE WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        month1_stock = 0
        month2_stock = 0
        close_stock = 0
        RQD_QTY = 0
        
        Set RSTRTRXFILE2 = New ADODB.Recordset
        RSTRTRXFILE2.Open "SELECT DISTINCT ITEM_CODE, P_RETAIL, BARCODE FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
        If RSTRTRXFILE2.RecordCount <= 1 Then
            RSTRTRXFILE2.Close
            Set RSTRTRXFILE2 = Nothing
            GoTo SKIP
        End If
        RSTRTRXFILE2.Close
        Set RSTRTRXFILE2 = Nothing
        
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE.EOF
            close_stock = close_stock + RSTRTRXFILE!BAL_QTY
            RSTRTRXFILE.MoveNext
        Loop
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
        Set RSTRTRXFILE = New ADODB.Recordset
        RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE.EOF
            month1_stock = month1_stock + RSTRTRXFILE!QTY
            RSTRTRXFILE.MoveNext
        Loop
        RSTRTRXFILE.Close
        Set RSTRTRXFILE = Nothing
        
        DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
        Set RSTRTRXFILE2 = New ADODB.Recordset
        RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND P_RETAIL = " & rststock!P_RETAIL & " AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE2.EOF
            month2_stock = month2_stock + RSTRTRXFILE2!QTY
            RSTRTRXFILE2.MoveNext
        Loop
        RSTRTRXFILE2.Close
        
        If close_stock < 0 Then
            RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
        Else
            RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
        End If
        If RQD_QTY < 0 Then RQD_QTY = "0"
        If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
        If OptStock.value = True And RQD_QTY <= 0 Then GoTo SKIP
        
        i = i + 1
        GRDSTOCK.Rows = GRDSTOCK.Rows + 1
        GRDSTOCK.FixedRows = 1
        GRDSTOCK.TextMatrix(i, 0) = i
        GRDSTOCK.TextMatrix(i, 1) = rststock!ITEM_CODE
        GRDSTOCK.TextMatrix(i, 2) = rststock!ITEM_NAME
        GRDSTOCK.TextMatrix(i, 3) = month2_stock
        GRDSTOCK.TextMatrix(i, 4) = close_stock
        If Val(GRDSTOCK.TextMatrix(i, 4)) < 0 Then
            GRDSTOCK.TextMatrix(i, 5) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
        Else
            GRDSTOCK.TextMatrix(i, 5) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
        End If
        GRDSTOCK.TextMatrix(i, 6) = IIf(IsNull(rststock!P_RETAIL), "", Format(rststock!P_RETAIL, "0.00"))
        If Val(GRDSTOCK.TextMatrix(i, 5)) < 0 Then GRDSTOCK.TextMatrix(i, 5) = "0"
        If Val(GRDSTOCK.TextMatrix(i, 4)) <= 0 And Val(GRDSTOCK.TextMatrix(i, 5)) = 0 Then GRDSTOCK.TextMatrix(i, 5) = "1"
        GRDSTOCK.TextMatrix(i, 7) = "Opening Stock"
        Set RSTSUPPLIER = New ADODB.Recordset
        RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
        If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
            GRDSTOCK.TextMatrix(i, 7) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
             GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
        Else
            Set RSTSUPPLIER2 = New ADODB.Recordset
            RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
            If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                GRDSTOCK.TextMatrix(i, 7) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                GRDSTOCK.TextMatrix(i, 8) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
            Else
                GRDSTOCK.TextMatrix(i, 8) = ""
            End If
            RSTSUPPLIER2.Close
            Set RSTSUPPLIER2 = Nothing
        End If
        RSTSUPPLIER.Close
        Set RSTSUPPLIER = Nothing
        'GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
SKIP:
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLTRXTOTAL.Caption = ""
    For i = 1 To GRDSTOCK.Rows - 1
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
    Next i
    
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Function

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Function Order4()
    Dim rststock, RSTRTRXFILE, RSTRTRXFILE2, RSTRTRXFILE3, RSTRTRXFILE4, RSTRTRXFILE5, RSTSUPPLIER, RSTSUPPLIER2 As ADODB.Recordset
    Dim i, COUNT As Integer
    Dim month1_stock, month2_stock, close_stock, RQD_QTY As Double
    'PHY_FLAG = True
    
    If Val(txtmonth2.Text) = 0 Or Val(txtmonth1.Text) = 0 Then
        MsgBox "Please enter proper values", vbOKOnly, "Order"
        txtmonth2.SetFocus
        Exit Function
    End If
    LBLTRXTOTAL.Caption = ""
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    'Screen.MousePointer = vbHourglass
    'GRDSTOCK.TextMatrix(0, 3) = "Sold last " & Val(txtmonth2.Text) & " days"
    'GRDSTOCK.TextMatrix(0, 5) = "Required for " & Val(txtmonth1.Text) & " days"
    
    Dim Item_found As Boolean
    Dim STOCKQTY As Double
    On Error GoTo Errhand
    Set rststock = New ADODB.Recordset
    If DataList1.BoundText = "" Then
        rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Else
        If optCategory.value = True Then
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND CATEGORY = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        Else
            rststock.Open "SELECT * FROM ITEMMAST WHERE ITEM_NAME Like '%" & Me.tXTMEDICINE.Text & "%' AND ITEM_NAME Like '%" & Me.TxtItem1.Text & "%'AND MANUFACTURER = '" & DataList1.BoundText & "' ORDER BY ITEM_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
        End If
    End If
    If rststock.RecordCount > 0 Then
        Screen.MousePointer = vbHourglass
        MDIMAIN.vbalProgressBar1.Visible = True
        MDIMAIN.vbalProgressBar1.value = 0
        MDIMAIN.vbalProgressBar1.ShowText = True
        MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
        MDIMAIN.vbalProgressBar1.Max = rststock.RecordCount
    End If
    Do Until rststock.EOF
        
        COUNT = 0
        Set RSTRTRXFILE4 = New ADODB.Recordset
        RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, BARCODE FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE4.EOF
            If RSTRTRXFILE4.RecordCount = 1 Then
                COUNT = 0
                Exit Do
            End If
            month1_stock = 0
            month2_stock = 0
            close_stock = 0
            RQD_QTY = 0
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                month1_stock = month1_stock + RSTRTRXFILE!QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
            Set RSTRTRXFILE2 = New ADODB.Recordset
            RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE2.EOF
                month2_stock = month2_stock + RSTRTRXFILE2!QTY
                RSTRTRXFILE2.MoveNext
            Loop
            RSTRTRXFILE2.Close
            
            If close_stock < 0 Then
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
            Else
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
            End If
            If RQD_QTY < 0 Then RQD_QTY = "0"
            If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            If RQD_QTY > 0 Then COUNT = COUNT + 1
            RSTRTRXFILE4.MoveNext
        Loop
        RSTRTRXFILE4.Close
        Set RSTRTRXFILE4 = Nothing
            
        Item_found = False
        Set RSTRTRXFILE4 = New ADODB.Recordset
        RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, BARCODE FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
        Do Until RSTRTRXFILE4.EOF
            month1_stock = 0
            month2_stock = 0
            close_stock = 0
            RQD_QTY = 0
            
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
            Set RSTRTRXFILE = New ADODB.Recordset
            RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE.EOF
                month1_stock = month1_stock + RSTRTRXFILE!QTY
                RSTRTRXFILE.MoveNext
            Loop
            RSTRTRXFILE.Close
            Set RSTRTRXFILE = Nothing
            
            DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
            Set RSTRTRXFILE2 = New ADODB.Recordset
            RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE2.EOF
                month2_stock = month2_stock + RSTRTRXFILE2!QTY
                RSTRTRXFILE2.MoveNext
            Loop
            RSTRTRXFILE2.Close
            
            If close_stock < 0 Then
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
            Else
                RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
            End If
            If RQD_QTY < 0 Then RQD_QTY = "0"
            If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            If OptStock.value = True And RQD_QTY <= 0 Then
                If COUNT = 0 Then GoTo SKIP2
            End If
            i = i + 1
            GRDSTOCK.Rows = GRDSTOCK.Rows + 1
            Item_found = True
            GRDSTOCK.FixedRows = 1
            With GRDSTOCK
                .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
                Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                .TextMatrix(i, 11) = "N"
                .TextMatrix(i, 1) = i
                .TextMatrix(i, 2) = rststock!ITEM_CODE
                .TextMatrix(i, 3) = rststock!ITEM_NAME
                '.TextMatrix(i, 3) = "==>" & rststock!ITEM_NAME
                .TextMatrix(i, 4) = month2_stock
                .TextMatrix(i, 5) = close_stock
                If Val(.TextMatrix(i, 5)) < 0 Then
                    .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                Else
                    .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                End If
                If COUNT = 0 Then .TextMatrix(i, 10) = IIf(IsNull(RSTRTRXFILE4!BARCODE), "", RSTRTRXFILE4!BARCODE)
                If Val(.TextMatrix(i, 6)) < 0 Then .TextMatrix(i, 6) = "0"
                If Val(.TextMatrix(i, 5)) <= 0 And Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = "1"
                .TextMatrix(i, 8) = "Opening Stock"
                Set RSTSUPPLIER = New ADODB.Recordset
                RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                    .TextMatrix(i, 8) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                    .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
                    .TextMatrix(i, 7) = IIf(IsNull(RSTSUPPLIER!P_RETAIL), "", Format(RSTSUPPLIER!P_RETAIL, "0.00"))
                Else
                    Set RSTSUPPLIER2 = New ADODB.Recordset
                    RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                        .TextMatrix(i, 8) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                        .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
                        .TextMatrix(i, 7) = IIf(IsNull(RSTSUPPLIER2!P_RETAIL), "", Format(RSTSUPPLIER2!P_RETAIL, "0.00"))
                    Else
                        .TextMatrix(i, 9) = ""
                    End If
                    RSTSUPPLIER2.Close
                    Set RSTSUPPLIER2 = Nothing
                End If
                RSTSUPPLIER.Close
                Set RSTSUPPLIER = Nothing
            End With
            LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
            Item_found = True
            Exit Do
SKIP2:
            RSTRTRXFILE4.MoveNext
        Loop
        RSTRTRXFILE4.Close
        Set RSTRTRXFILE4 = Nothing
        
        If COUNT > 0 Then
            Set RSTRTRXFILE4 = New ADODB.Recordset
            RSTRTRXFILE4.Open "SELECT DISTINCT ITEM_CODE, BARCODE FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
            Do Until RSTRTRXFILE4.EOF
                If RSTRTRXFILE4.RecordCount = 1 Then Exit Do
                month1_stock = 0
                month2_stock = 0
                close_stock = 0
                RQD_QTY = 0
                
                Set RSTRTRXFILE = New ADODB.Recordset
                RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND BAL_QTY > 0", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE.EOF
                    close_stock = close_stock + RSTRTRXFILE!BAL_QTY
                    RSTRTRXFILE.MoveNext
                Loop
                RSTRTRXFILE.Close
                Set RSTRTRXFILE = Nothing
                
                DTFROM.value = DateAdd("d", -Val(txtmonth1.Text), Date)
                Set RSTRTRXFILE = New ADODB.Recordset
                RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE.EOF
                    month1_stock = month1_stock + RSTRTRXFILE!QTY
                    RSTRTRXFILE.MoveNext
                Loop
                RSTRTRXFILE.Close
                Set RSTRTRXFILE = Nothing
                
                DTFROM.value = DateAdd("d", -Val(txtmonth2.Text), Date)
                Set RSTRTRXFILE2 = New ADODB.Recordset
                RSTRTRXFILE2.Open "SELECT * FROM TRXFILE WHERE ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' AND VCH_DATE >='" & Format(DTFROM.value, "yyyy/mm/dd") & "'", db, adOpenStatic, adLockReadOnly
                Do Until RSTRTRXFILE2.EOF
                    month2_stock = month2_stock + RSTRTRXFILE2!QTY
                    RSTRTRXFILE2.MoveNext
                Loop
                RSTRTRXFILE2.Close
                
                If close_stock < 0 Then
                    RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                Else
                    RQD_QTY = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                End If
                If RQD_QTY < 0 Then RQD_QTY = "0"
                If close_stock <= 0 And RQD_QTY = 0 Then RQD_QTY = 1
            
                i = i + 1
                GRDSTOCK.Rows = GRDSTOCK.Rows + 1
                Item_found = True
                GRDSTOCK.FixedRows = 1
                With GRDSTOCK
                    .Row = i: .Col = 0: .CellPictureAlignment = 4 ' Align the checkbox
                    Set .CellPicture = picChecked.Picture  ' Set the default checkbox picture to the empty box
                    .TextMatrix(i, 11) = "N"
                    .TextMatrix(i, 1) = i
                    .TextMatrix(i, 2) = rststock!ITEM_CODE
                    .TextMatrix(i, 3) = "==>" & rststock!ITEM_NAME
                    .TextMatrix(i, 4) = month2_stock
                    .TextMatrix(i, 5) = close_stock
                    If Val(.TextMatrix(i, 5)) < 0 Then
                        .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)), 0)
                    Else
                        .TextMatrix(i, 6) = Round((month2_stock / Val(txtmonth2.Text) * Val(txtmonth1.Text)) - close_stock, 0)
                    End If
                    .TextMatrix(i, 10) = IIf(IsNull(RSTRTRXFILE4!BARCODE), "", RSTRTRXFILE4!BARCODE)
                    If Val(.TextMatrix(i, 6)) < 0 Then .TextMatrix(i, 6) = "0"
                    If Val(.TextMatrix(i, 5)) <= 0 And Val(.TextMatrix(i, 6)) = 0 Then .TextMatrix(i, 6) = "1"
                    .TextMatrix(i, 8) = "Opening Stock"
                    Set RSTSUPPLIER = New ADODB.Recordset
                    RSTSUPPLIER.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PI' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                    If Not (RSTSUPPLIER.EOF And RSTSUPPLIER.BOF) Then
                        .TextMatrix(i, 8) = "P- " & RSTSUPPLIER!VCH_NO & IIf(IsNull(RSTSUPPLIER!VCH_DESC), "", ", " & Mid(RSTSUPPLIER!VCH_DESC, 15))
                        .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER!ITEM_COST), "", Format(RSTSUPPLIER!ITEM_COST, "0.00"))
                        .TextMatrix(i, 7) = IIf(IsNull(RSTSUPPLIER!P_RETAIL), "", Format(RSTSUPPLIER!P_RETAIL, "0.00"))
                    Else
                        Set RSTSUPPLIER2 = New ADODB.Recordset
                        RSTSUPPLIER2.Open "SELECT * FROM RTRXFILE WHERE TRX_TYPE = 'PW' AND ITEM_CODE = '" & RSTRTRXFILE4!ITEM_CODE & "' AND BARCODE = '" & RSTRTRXFILE4!BARCODE & "' ORDER BY  VCH_NO DESC", db, adOpenStatic, adLockReadOnly
                        If Not (RSTSUPPLIER2.EOF And RSTSUPPLIER2.BOF) Then
                            .TextMatrix(i, 8) = "W- " & RSTSUPPLIER2!VCH_NO & IIf(IsNull(RSTSUPPLIER2!VCH_DESC), "", ", " & Mid(RSTSUPPLIER2!VCH_DESC, 15))
                            .TextMatrix(i, 9) = IIf(IsNull(RSTSUPPLIER2!ITEM_COST), "", Format(RSTSUPPLIER2!ITEM_COST, "0.00"))
                            .TextMatrix(i, 7) = IIf(IsNull(RSTSUPPLIER2!P_RETAIL), "", Format(RSTSUPPLIER2!P_RETAIL, "0.00"))
                        Else
                            .TextMatrix(i, 9) = ""
                        End If
                        RSTSUPPLIER2.Close
                        Set RSTSUPPLIER2 = Nothing
                    End If
                    RSTSUPPLIER.Close
                    Set RSTSUPPLIER = Nothing
                End With
                RSTRTRXFILE4.MoveNext
            Loop
            RSTRTRXFILE4.Close
            Set RSTRTRXFILE4 = Nothing
        End If
        'GRDSTOCK.TextMatrix(i, 9) = IIf(IsNull(rststock!BARCODE), "", rststock!BARCODE)
        
        MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    LBLTRXTOTAL.Caption = ""
    For i = 1 To GRDSTOCK.Rows - 1
        LBLTRXTOTAL.Caption = Val(LBLTRXTOTAL.Caption) + (Val(GRDSTOCK.TextMatrix(i, 9)) * Val(GRDSTOCK.TextMatrix(i, 6)))
    Next i
    
    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Function

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Function

Private Sub txtmonth1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtmonth1.Text) = 0 Then Exit Sub
            tXTMEDICINE.SetFocus
        Case vbKeyEscape
            txtmonth2.SetFocus
    End Select
End Sub

Private Sub txtmonth1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtmonth2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(txtmonth2.Text) = 0 Then Exit Sub
            txtmonth1.SetFocus
    End Select
End Sub

Private Sub txtmonth2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub TXTsample_GotFocus()
    TXTsample.SelStart = 0
    TXTsample.SelLength = Len(TXTsample.Text)
End Sub

Private Sub TXTsample_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Select Case GRDSTOCK.Col
                Case 3, 4, 5, 6
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Val(TXTsample.Text)
                Case 7
                    GRDSTOCK.TextMatrix(GRDSTOCK.Row, GRDSTOCK.Col) = Trim(TXTsample.Text)
            End Select
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
        Case vbKeyEscape
            TXTsample.Visible = False
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub TXTsample_KeyPress(KeyAscii As Integer)
    Select Case GRDSTOCK.Col
        Case 3, 4, 5, 6
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
        Case 7
             Select Case KeyAscii
                Case Asc("'"), Asc("["), Asc("]"), Asc("\")
                    KeyAscii = 0
                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                Case Else
                    KeyAscii = 0
            End Select
'        Case 7
'             Select Case KeyAscii
'                Asc("'"), Asc("["), Asc("]"), Asc("\")
'                    KeyAscii = 0
'                Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, Asc(".")
'                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'                Case Else
'                    KeyAscii = 0
'            End Select
    End Select
End Sub

Private Sub TxtItem1_GotFocus()
    TxtItem1.SelStart = 0
    TxtItem1.SelLength = Len(tXTMEDICINE.Text)
    'Call Fillgrid
End Sub

Private Sub TxtItem1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Call CMDDISPLAY_Click
    End Select

End Sub

