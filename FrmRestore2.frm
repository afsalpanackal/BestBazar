VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRestore 
   Caption         =   "Restore Database"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "FrmRestore2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   6510
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdExit 
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
      Height          =   495
      Left            =   5160
      TabIndex        =   3
      Top             =   1575
      Width           =   1245
   End
   Begin VB.CommandButton CmdLoad 
      Caption         =   "&Load"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   2
      Top             =   1005
      Width           =   1245
   End
   Begin VB.TextBox txtPath 
      Height          =   330
      Left            =   135
      TabIndex        =   1
      Top             =   540
      Width           =   4905
   End
   Begin VB.CommandButton CmdBrowse 
      Caption         =   "&Browse"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   450
      Width           =   1245
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdBrowse_Click()
    On Error GoTo errhandler
    CommonDialog1.CancelError = True
    CommonDialog1.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist + cdlOFNFileMustExist
    CommonDialog1.Filter = "All Files (*.*)"
    CommonDialog1.ShowOpen
    txtPath.Text = CommonDialog1.FileName
    Exit Sub
errhandler:
    MsgBox Err.Description
End Sub

Private Sub CmdExit_Click()
    Unload Me
    End
End Sub

Private Sub CmdLoad_Click()
    On Error GoTo errhandler
    FileLen (txtPath.Text)
    If (MsgBox("Are you sure you want to Restore the Database?", vbYesNo, "LOGIN") = vbNo) Then Exit Sub
    Dim SourceFile, DestinationFile, tryagain, result
    Dim strBackupEXT As String
    On Error GoTo errhandler
    
    Screen.MousePointer = vbHourglass
    strBackupEXT = "BK" & Format(Format(Date, "ddmmyy"), "000000") & Format(Format(Time, "hhmmss"), "000000")
    Screen.MousePointer = vbHourglass
    SourceFile = App.Path & "\SOFT.SML"
    DestinationFile = App.Path & "\" & strBackupEXT
    result = apiCopyFile(SourceFile, DestinationFile, False)
    
    SourceFile = CommonDialog1.FileName
    DestinationFile = App.Path & "\SOFT.SML"
    result = apiCopyFile(SourceFile, DestinationFile, False)
    
    Screen.MousePointer = vbNormal
    If result = 0 Then
        MsgBox "Restore Failed", vbOKOnly, "RESTORE"
    Else
        MsgBox "Successfully Restored", vbOKOnly, "RESTORE"
    End If
    Exit Sub
errhandler:
    Screen.MousePointer = vbNormal
    If Err.Number = 53 Then
        MsgBox "File not exists..", vbOKOnly, "RESTORE"
    Else
        MsgBox Err.Description
    End If
End Sub
