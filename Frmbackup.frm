VERSION 5.00
Begin VB.Form fRMbackup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Backup"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   Icon            =   "Frmbackup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4125
   Begin VB.DriveListBox DrvDest 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   375
      TabIndex        =   4
      Top             =   510
      Width           =   3330
   End
   Begin VB.DirListBox DirDstn 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   975
      Width           =   3315
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1995
      TabIndex        =   1
      Top             =   3390
      Width           =   1695
   End
   Begin VB.CommandButton CMDBakup 
      Caption         =   "&Backup"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   330
      TabIndex        =   0
      Top             =   3390
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select Destination Path"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Index           =   6
      Left            =   345
      TabIndex        =   3
      Top             =   60
      Width           =   3375
   End
End
Attribute VB_Name = "fRMbackup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Private Sub CMDBakup_Click()

    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    MDIMAIN.vbalProgressBar1.text = "PLEASE WAIT..."
            
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTBALQTY As ADODB.Recordset
    Dim INWARD As Double
    Dim OUTWARD As Double
    Dim BALQTY As Double
    Dim DIFFQTY As Double
    Dim i As Long
    
    
    Screen.MousePointer = vbHourglass
    On Error GoTo handler
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
    Do Until RSTITEMMAST.EOF
        
        INWARD = 0
        OUTWARD = 0
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(QTY + FREE_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
                    
        Set rststock = New ADODB.Recordset
        rststock.Open "Select SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rststock.EOF And rststock.BOF) Then
            OUTWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
    
        RSTITEMMAST!CLOSE_QTY = Round(INWARD - OUTWARD, 2)
        RSTITEMMAST!RCPT_QTY = INWARD
        RSTITEMMAST!ISSUE_QTY = OUTWARD
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
            
    'On Error GoTo handler
    Dim cmd As String
    Dim strBackupEXT As String
    If Not FileExists(App.Path & "\mysqldump.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    strBackupEXT = "bk" & Format(Format(Date, "ddmmyy"), "000000") & Format(Format(Time, "HHMMSS"), "")
    DoEvents
    'cmd = Chr(34) & "C:\wamp\bin\mysql\mysql5.5.8\bin\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret --routines --comments " & dbase1 & " > " & DirDstn & "\" & strBackupEXT
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " > " & DirDstn & "\" & strBackupEXT
    Call execCommand(cmd)
    
    err.Clear
    
    Screen.MousePointer = vbNormal
    MDIMAIN.vbalProgressBar1.text = "Successfully Completed..."
    MsgBox "Back-up complete !!", vbOKOnly, "Back Up!!!!"
    MDIMAIN.vbalProgressBar1.Visible = False
    Exit Sub
    
handler:
    Screen.MousePointer = vbNormal
    Select Case err.Number
        Case 70
        MsgBox "Error No. > " & err.Number & " / " & err.Description
        Resume Next
        Case 75
        Resume Next
        Case Else
        MsgBox "Error No. > " & err.Number & " / " & err.Description
    End Select


End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub DrvDest_Change()
    On Error GoTo ErrHand
    DirDstn.Path = DrvDest
    Exit Sub
ErrHand:
    If err.Number = 68 Then
        DrvDest = "C:\"
        DirDstn.Path = "C:\"
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub Form_Load()
    cetre Me
End Sub

Private Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long
 
    cmd = "cmd /c " & cmd
    result = Shell(cmd, vbHide)
     
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

