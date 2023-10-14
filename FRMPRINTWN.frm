VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FRMPRINTEXP 
   ClientHeight    =   1455
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   7440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   1455
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FRMEPRINT 
      BackColor       =   &H00C0C0FF&
      Height          =   1440
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   7455
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         Height          =   405
         Left            =   6060
         TabIndex        =   5
         Top             =   855
         Width           =   1215
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "&OK"
         Height          =   405
         Left            =   4770
         TabIndex        =   4
         Top             =   855
         Width           =   1200
      End
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
         Left            =   5445
         TabIndex        =   1
         Top             =   195
         Width           =   1890
      End
      Begin MSDataListLib.DataCombo CMBexpdist 
         Height          =   330
         Left            =   1200
         TabIndex        =   2
         Top             =   210
         Width           =   3495
         _ExtentX        =   6165
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
         Caption         =   "SUPPLIER"
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
         Left            =   165
         TabIndex        =   3
         Top             =   255
         Width           =   960
      End
   End
End
Attribute VB_Name = "FRMPRINTEXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ACT_REC As New ADODB.Recordset
Dim ACT_FLAG As Boolean

Private Sub chksent_Click()
    Dim RSTFLAG As ADODB.Recordset
    Dim i, n As Integer
    
    If Trim(CMBexpdist.Text) = "" Then
        chksent.Value = 0
        Exit Sub
    End If
    
    On Error GoTo ErrHand
    
    n = 0
    If chksent.Value = 1 Then
        If (MsgBox("ARE YOU SURE YOU WANT TO MARK " & Trim(CMBexpdist.Text) & " AS SENT", vbYesNo) = vbNo) Then
            chksent.Value = 0
            Exit Sub
        End If
        For i = 0 To FRMWARRANTY.grdcount.rows - 1
            If FRMWARRANTY.grdcount.TextMatrix(i, 8) = CMBexpdist.Text Then
                Set RSTFLAG = New ADODB.Recordset
                RSTFLAG.Open "SELECT * from WAR_TRXFILE where VCH_NO = " & Val(FRMWARRANTY.grdcount.TextMatrix(i, 2)) & " AND LINE_NO = " & Val(FRMWARRANTY.grdcount.TextMatrix(i, 14)) & " ", db, adOpenStatic, adLockOptimistic, adCmdText
                If Not (RSTFLAG.EOF And RSTFLAG.BOF) Then
                    n = n + 1
                    RSTFLAG!check_flag = "Y"
                    RSTFLAG!SENT_DATE = Date
                    RSTFLAG.Update
                End If
                RSTFLAG.Close
                Set RSTFLAG = Nothing
            End If
        Next i
        Call fillcombo
        MsgBox n & " Items Marked as Sent Successfully", vbOKOnly, "Warranty Replacement"
        CMBexpdist.Text = ""
        chksent.Value = 0
    End If
    Screen.MousePointer = vbNormal
    Exit Sub

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description

End Sub

Private Sub CmdExit_Click()
    Unload Me
    MDIMAIN.Enabled = True
    FRMWARRANTY.Enabled = True
    Call FRMWARRANTY.Fillgrid
    FRMWARRANTY.SetFocus
End Sub

Private Sub cmdOK_Click()
    If CMBexpdist.BoundText = "" Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ErrHand
    
    Call cmdReportGenerate_Click
    
    Open Rptpath & "repo.bat" For Output As #1 '//Creating Batch file

    Print #1, "COPY/B " & Rptpath & "Report.PRN " & DMPrint
    Print #1, "EXIT"
    Close #1

    '//HERE write the proper path where your command.com file exist
    Shell "C:\WINDOWS\SYSTEM32\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    'Shell "C:\WINDOW\COMMAND.COM /C " & Rptpath & "REPO.BAT N", vbHide
    Screen.MousePointer = vbNormal
    Exit Sub
ErrHand:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub Form_Load()
    ACT_FLAG = True
    cetre Me
    fillcombo
End Sub

Private Sub cmdReportGenerate_Click()
    Dim RSTEXP As ADODB.Recordset
    Dim RSTCOMPANY As ADODB.Recordset
    Dim vlineCount As Integer
    Dim vpageCount As Integer
    Dim TOTAL As Double
    Dim i As Long
    
    vlineCount = 0
    vpageCount = 1
    SN = 0
    
    'Set Rs = New ADODB.Recordset
    'strSQL = "Select * from products"
    'Rs.Open strSQL, cnn
    
    '//NOTE : Report file name should never contain blank space.
    Open Rptpath & "Report.PRN" For Output As #1 '//Report file Creation
    
    '//Create Report Heading
    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
    
    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
            Chr(106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
    Print #1, Chr(27) & Chr(72)
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenForwardOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        Print #1, Chr(27) & Chr(71) & Space(22) & Chr(14) & Chr(15) & Space(1) & "EXPIRY RETURN"
        Print #1, Chr(27) & Chr(71) & Chr(10) & Space(6) & " Name of Retailer:" & _
              Space(7) & Chr(14) & Chr(15) & AlignLeft(RSTCOMPANY!COMP_NAME, 30) & _
              Chr(27) & Chr(72)
        Print #1, Space(32) & AlignLeft(RSTCOMPANY!Address & ", " & RSTCOMPANY!HO_NAME, 140)
        Print #1, Space(48) & AlignRight("DL NO. " & RSTCOMPANY!CST, 25)
        Print #1, Space(48) & AlignRight(RSTCOMPANY!DL_NO, 25)
        Print #1, Space(48) & AlignRight("TIN No. " & RSTCOMPANY!KGST, 25)
        Print #1,
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing

   ' Print #1, Space(7) & "Pandalam 688006" & Space(15) & "DL No. 6-176/20/2003 Dtd. 31.10.2003"
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
    RSTEXP.Open "SELECT * From EXPLIST WHERE EX_DISTI = '" & Trim(CMBexpdist.Text) & "' AND EX_FLAG ='N'", db, adOpenForwardOnly
    Do Until RSTEXP.EOF
            i = i + 1
         Print #1, Space(7) & AlignRight(str(i), 3) & Space(2) & _
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

    Dim x As Integer
    
    For x = 1 To vSpace
        RepeatString = RepeatString & vStr
    Next x
End Function


Private Sub Form_Unload(Cancel As Integer)
    ACT_REC.Close
End Sub

Private Function fillcombo()
    On Error GoTo ErrHand
    
    Screen.MousePointer = vbHourglass
    Set CMBexpdist.DataSource = Nothing
    If ACT_FLAG = True Then
        ACT_REC.Open "select Distinct DIST_NAME from WAR_TRXFILE WHERE CHECK_FLAG ='N' AND DIST_NAME <> '' ORDER BY DIST_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    Else
        ACT_REC.Close
        ACT_REC.Open "select Distinct DIST_NAME from WAR_TRXFILE WHERE CHECK_FLAG ='N' AND DIST_NAME <> '' ORDER BY DIST_NAME", db, adOpenStatic, adLockReadOnly, adCmdText
        ACT_FLAG = False
    End If
    
    Set CMBexpdist.RowSource = ACT_REC
    CMBexpdist.ListField = "DIST_NAME"
    CMBexpdist.BoundColumn = "DIST_NAME"
    Screen.MousePointer = vbNormal
    Exit Function

ErrHand:
    Screen.MousePointer = vbNormal
     MsgBox err.Description
End Function
