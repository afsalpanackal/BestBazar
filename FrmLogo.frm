VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fRMLOGO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Picture"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "FrmLogo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   4230
   Begin VB.CommandButton CmdEXIT 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1410
      TabIndex        =   3
      Top             =   3090
      Width           =   1485
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "Reset"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1410
      TabIndex        =   2
      Top             =   1575
      Width           =   1485
   End
   Begin VB.CommandButton cmddelphoto 
      Caption         =   "Remove Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1410
      TabIndex        =   1
      Top             =   960
      Width           =   1485
   End
   Begin VB.CommandButton CMDBROWSE 
      Caption         =   "Load Picture"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1425
      TabIndex        =   0
      Top             =   345
      Width           =   1470
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "fRMLOGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDBROWSE_Click()
    
    On Error GoTo errhandler
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "Picture Files (*.jpg)|*.jpg"
    CommonDialog1.ShowOpen
    MDIMAIN.Picture = LoadPicture(CommonDialog1.FileName)
    
    Dim RSTITEMMAST As ADODB.Recordset
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM COMPINFO WHERE COMP_CODE='001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!COMP_LOGO = CommonDialog1.FileName
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MDIMAIN.Hide
    MDIMAIN.Show
    
    Exit Sub
errhandler:
    Select Case Err
    Case 32755 '  Dialog Cancelled
        'MsgBox "you cancelled the dialog box"
    Case Else
        MsgBox "Unexpected error. Err " & Err & " : " & Error
    End Select
    
End Sub

Private Sub cmddelphoto_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    On Error GoTo errhandler
    CommonDialog1.FileName = ""
    MDIMAIN.Picture = LoadPicture("")
    
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM COMPINFO WHERE COMP_CODE='001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
        RSTITEMMAST!COMP_LOGO = CommonDialog1.FileName
        RSTITEMMAST.Update
    End If
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    MDIMAIN.Hide
    MDIMAIN.Show
    Exit Sub
errhandler:
    MsgBox "Unexpected error. Err " & Err & " : " & Error
End Sub


Private Sub CmdEXIT_Click()
    Unload Me
End Sub

Private Sub CMDRESET_Click()
    Dim RSTCOMPANY As ADODB.Recordset
    On Error GoTo errhandler
    db.Execute "Update COMPINFO set COMP_LOGO = Null WHERE COMP_CODE='001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.value) & "'"
    MDIMAIN.Picture = LoadPicture("")
    Exit Sub
errhandler:
    MsgBox "Unexpected error. Err " & Err & " : " & Error
End Sub

