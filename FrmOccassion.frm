VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmOccassion 
   BackColor       =   &H00CFEFE4&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminder"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   17655
   Icon            =   "FrmOccassion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   17655
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
      Left            =   15
      TabIndex        =   6
      Top             =   8190
      Visible         =   0   'False
      Width           =   1410
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00CFEFE4&
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   -75
      Width           =   14835
      Begin VB.TextBox Txtdays 
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
         Left            =   1815
         MaxLength       =   2
         TabIndex        =   7
         Top             =   180
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   390
         Left            =   10515
         TabIndex        =   8
         Top             =   45
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
         Format          =   110362627
         CurrentDate     =   40498
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CFEFE4&
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
         Left            =   2985
         TabIndex        =   5
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00CFEFE4&
         Caption         =   "Occassions within"
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
         Left            =   150
         TabIndex        =   4
         Top             =   240
         Width           =   1635
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
      TabIndex        =   1
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
      Left            =   12315
      TabIndex        =   0
      Top             =   8190
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid GRDSTOCK 
      Height          =   7560
      Left            =   0
      TabIndex        =   2
      Top             =   585
      Width           =   17580
      _ExtentX        =   31009
      _ExtentY        =   13335
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
End
Attribute VB_Name = "FrmOccassion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PHY_REC As New ADODB.Recordset
Dim PHY_FLAG As Boolean

Private Sub CMDDETAILS_Click()
    Dim RSTTEM As ADODB.Recordset
    Dim i As Long
    On Error GoTo Errhand
    db.Execute "DELETE * FROM TEMPSTK"
    
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "ITEM CODE"
    GRDSTOCK.TextMatrix(0, 2) = "ITEM NAME"
    GRDSTOCK.TextMatrix(0, 3) = "SOLD"
    GRDSTOCK.TextMatrix(0, 4) = "SHELF"
    GRDSTOCK.TextMatrix(0, 5) = "RQD QTY"
    GRDSTOCK.TextMatrix(0, 6) = "MRP"
    GRDSTOCK.TextMatrix(0, 7) = "SUPPLIER"
    
    If GRDSTOCK.Rows <= 1 Then Exit Sub
    Set RSTTEM = New ADODB.Recordset
    RSTTEM.Open "SELECT * FROM TEMPSTK", db, adOpenStatic, adLockOptimistic, adCmdText
    For i = 1 To GRDSTOCK.Rows - 1
        RSTTEM.AddNew
        RSTTEM!ITEM_CODE = GRDSTOCK.TextMatrix(i, 1)
        RSTTEM!ITEM_NAME = GRDSTOCK.TextMatrix(i, 2)
        RSTTEM!INQTY = GRDSTOCK.TextMatrix(i, 3)
        RSTTEM!OUTQTY = GRDSTOCK.TextMatrix(i, 4)
        RSTTEM!CLOSE_QTY = GRDSTOCK.TextMatrix(i, 5)
        RSTTEM!CLOSE_VAL = GRDSTOCK.TextMatrix(i, 6)
        RSTTEM!DIFF_QTY = Trim(GRDSTOCK.TextMatrix(i, 7))
        'RSTTEM!OPQTY = 0 'GRDSTOCK.TextMatrix(i, 3)
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
    
    frmreport.Caption = "STOCK REPORT"
    ReportNameVar = "D:\EzBiz\RptReport.RPT"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
    Set CRXFormulaFields = Report.FormulaFields
    
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables(i).SetLogOnInfo "ConnectionName", "C:\Users\Public\WINSYS.SYS", "admin", "###DATABASE%%%RET"
    Next i
    Report.OpenSubreport("RptReport.RPT").DiscardSavedData
    Report.OpenSubreport("RptReport.RPT").VerifyOnEveryPrint = True
    For Each CRXFormulaField In CRXFormulaFields
        'If CRXFormulaField.Name = "{@PERIOD}" Then CRXFormulaField.Text = "'" & DTFROM.value & "' & ' TO ' &'" & DTTO.value & "'"
        If CRXFormulaField.Name = "{@Head}" Then CRXFormulaField.Text = "'MY SHOP, ALAPPUZHA'"
    Next

    Call GENERATEREPORT
    
    Exit Sub
Errhand:
    MsgBox Err.Description
End Sub

Private Sub CMDDISPLAY_Click()
    Dim rststock As ADODB.Recordset
    Dim i As Long
    Dim DDate As Date
    On Error GoTo Errhand
    
    If Val(Txtdays.Text) = 0 Then
        MsgBox "Please enter the no. of days", vbOKOnly, "Reminder"
        Txtdays.SetFocus
        Exit Sub
    End If
    DTFROM.value = DateAdd("d", Val(Txtdays.Text), Date)
    GRDSTOCK.FixedRows = 0
    GRDSTOCK.Rows = 1
    i = 0
    Dim remdate  As String
    Screen.MousePointer = vbHourglass
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND [DOM] <=# " & Format(DTFROM.value, "MM,DD") & " # AND [DOM] >=# " & Format(Date, "MM,DD") & " # ORDER BY ACT_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    rststock.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' ORDER BY ACT_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        If (IsDate(rststock!DOM)) Then
            'DDate = Day(rststock!DOM) & "/" & Month(rststock!DOM) & "/" & Year(Date)
            remdate = Format(Day(rststock!DOM), "dd") & "/" & Format(Month(rststock!DOM), "00") & "/" & Format(Year(Date), "0000")
            If IsDate(remdate) Then DDate = remdate
            If DDate <= DTFROM.value And DDate >= Date Then
                i = i + 1
                GRDSTOCK.Rows = GRDSTOCK.Rows + 1
                GRDSTOCK.FixedRows = 1
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ACT_NAME
                GRDSTOCK.TextMatrix(i, 2) = rststock!Address
                GRDSTOCK.TextMatrix(i, 3) = rststock!TELNO
                GRDSTOCK.TextMatrix(i, 4) = rststock!FAXNO
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsDate(rststock!DOM), Format(rststock!DOM, "DD/MM/YYYY"), "")
                GRDSTOCK.TextMatrix(i, 6) = "Wedding Anniversary"
                GRDSTOCK.TextMatrix(i, 7) = DateDiff("d", Date, DDate)
                MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
            End If
        End If
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing
    
    Set rststock = New ADODB.Recordset
    'rststock.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' AND [DOM] <=# " & Format(DTFROM.value, "MM,DD") & " # AND [DOM] >=# " & Format(Date, "MM,DD") & " # ORDER BY ACT_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    rststock.Open "SELECT * FROM CUSTMAST WHERE ACT_CODE <> '130000' ORDER BY ACT_NAME ", db, adOpenStatic, adLockReadOnly, adCmdText
    Do Until rststock.EOF
        If (IsDate(rststock!DOB)) Then
            remdate = Format(Day(rststock!DOB), "dd") & "/" & Format(Month(rststock!DOB), "00") & "/" & Format(Year(Date), "0000")
            If IsDate(remdate) Then DDate = remdate
            If DDate <= DTFROM.value And DDate >= Date Then
                i = i + 1
                GRDSTOCK.Rows = GRDSTOCK.Rows + 1
                GRDSTOCK.FixedRows = 1
                GRDSTOCK.TextMatrix(i, 0) = i
                GRDSTOCK.TextMatrix(i, 1) = rststock!ACT_NAME
                GRDSTOCK.TextMatrix(i, 2) = rststock!Address
                GRDSTOCK.TextMatrix(i, 3) = rststock!TELNO
                GRDSTOCK.TextMatrix(i, 4) = rststock!FAXNO
                GRDSTOCK.TextMatrix(i, 5) = IIf(IsDate(rststock!DOB), Format(rststock!DOB, "DD/MM"), "")
                GRDSTOCK.TextMatrix(i, 6) = "Birthday"
                GRDSTOCK.TextMatrix(i, 7) = DateDiff("d", Date, DDate)
                MDIMAIN.vbalProgressBar1.value = MDIMAIN.vbalProgressBar1.value + 1
            End If
        End If
        rststock.MoveNext
    Loop
    rststock.Close
    Set rststock = Nothing


    MDIMAIN.vbalProgressBar1.Visible = False
    Screen.MousePointer = vbNormal
    Exit Sub

Errhand:
    Screen.MousePointer = vbNormal
     MsgBox Err.Description
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    GRDSTOCK.TextMatrix(0, 0) = "SL"
    GRDSTOCK.TextMatrix(0, 1) = "CUSTOMER"
    GRDSTOCK.TextMatrix(0, 2) = "Address"
    GRDSTOCK.TextMatrix(0, 3) = "Phone"
    GRDSTOCK.TextMatrix(0, 4) = "Mobile"
    GRDSTOCK.TextMatrix(0, 5) = "Date"
    GRDSTOCK.TextMatrix(0, 6) = "Event"
    GRDSTOCK.TextMatrix(0, 7) = "Days Left"
    
    GRDSTOCK.ColWidth(0) = 900
    GRDSTOCK.ColWidth(1) = 3000
    GRDSTOCK.ColWidth(2) = 6000
    GRDSTOCK.ColWidth(3) = 1300
    GRDSTOCK.ColWidth(4) = 1300
    GRDSTOCK.ColWidth(5) = 1300
    GRDSTOCK.ColWidth(6) = 2100
    GRDSTOCK.ColWidth(7) = 1100
    
    GRDSTOCK.ColAlignment(0) = 1
    GRDSTOCK.ColAlignment(1) = 1
    GRDSTOCK.ColAlignment(2) = 1
    GRDSTOCK.ColAlignment(3) = 4
    GRDSTOCK.ColAlignment(4) = 4
    GRDSTOCK.ColAlignment(5) = 4
    GRDSTOCK.ColAlignment(6) = 4
    GRDSTOCK.ColAlignment(7) = 4
    
    Txtdays.Text = 10
    PHY_FLAG = True
    DTFROM.value = "01/" & Month(Date) & "/" & Year(Date)
    Left = 500
    Top = 0
    Call CMDDISPLAY_Click
    'Height = 10000
    'Width = 12840
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PHY_FLAG = False Then PHY_REC.Close
   'Cancel = CLOSEALL
End Sub

Private Sub GRDSTOCK_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sitem As String
    Dim i As Long
    If GRDSTOCK.Rows = 1 Then Exit Sub
    Select Case KeyCode
        Case 114
            sitem = UCase(InputBox("Item Name..?", "ZERO STOCK"))
            For i = 1 To GRDSTOCK.Rows - 1
                    If Mid(GRDSTOCK.TextMatrix(i, 2), 1, Len(sitem)) = sitem Then
                        GRDSTOCK.Row = i
                        GRDSTOCK.TopRow = i
                    Exit For
                End If
            Next i
            GRDSTOCK.SetFocus
    End Select
End Sub

Private Sub Txtdays_GotFocus()
    Txtdays.SelStart = 0
    Txtdays.SelLength = Len(Txtdays.Text)
End Sub

Private Sub Txtdays_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If Val(Txtdays.Text) = 0 Then Exit Sub
            Call CMDDISPLAY_Click
    End Select
End Sub

Private Sub Txtdays_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
