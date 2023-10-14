VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FrmRestore 
   Caption         =   "Restore Database"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6510
   Icon            =   "FrmRestore.frx":0000
   LinkTopic       =   "Form1"
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
      Left            =   5175
      TabIndex        =   2
      Top             =   1020
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
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Private Sub CMDBROWSE_Click()
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

Private Sub cmdexit_Click()
    Unload Me
    End
End Sub

Private Sub CmdLoad_Click()
    On Error GoTo errhandler
    
    If Not FileExists(App.Path & "\mysql.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Sub
    End If
            
    FileLen (txtPath.Text)
    If (MsgBox("Are you sure you want to Restore the Database?", vbYesNo + vbDefaultButton2, "LOGIN") = vbNo) Then Exit Sub
    
    Dim cmd As String
    Screen.MousePointer = vbHourglass
'    DoEvents
'    cmd = Chr(34) & "C:\wamp\bin\mysql\mysql5.5.8\bin\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret create database" & dbase1 & " ;"
'    Call execCommand(cmd)
    
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret " & dbase1 & " <" & txtPath.Text
    Call execCommand(cmd)
 
    Screen.MousePointer = vbDefault
    MsgBox "Succesfully Restored", vbOKOnly, "RESTORE"
    
    Exit Sub
errhandler:
    Screen.MousePointer = vbNormal
    If Err.Number = 53 Then
        MsgBox "File not exists..", vbOKOnly, "RESTORE"
    Else
        MsgBox Err.Description
    End If
End Sub

Private Sub cmdRestore_Click()
    Screen.MousePointer = vbHourglass
    DoEvents
    Dim cmd As String
    If Not FileExists(App.Path & "\mysqldump.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Sub
    End If
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret " & dbase1 & " < c:\MyBackup.sql"
    Call execCommand(cmd)
 
    Screen.MousePointer = vbDefault
    MsgBox "done"
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

