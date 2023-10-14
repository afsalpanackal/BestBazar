VERSION 5.00
Begin VB.Form fRMPluUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update PLU Table"
   ClientHeight    =   4020
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   Icon            =   "Frmpluload.frx":0000
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
      Caption         =   "&Update"
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
Attribute VB_Name = "fRMPluUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CMDBakup_Click()
    
    Dim slcount As Long
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    If FileExists(App.Path & "\PLUPATH") Then
        Kill (App.Path & "\PLUPATH")
    End If
    Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
    Set ObjFile = ObjFso.CreateTextFile(App.Path & "\PLUPATH")
    ObjFile.WriteLine DirDstn
   
    
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    Dim RSTCOMPANY As ADODB.Recordset
    
    On Error GoTo CLOSEFILE
    Open DirDstn & "\PLU.txt" For Output As #1 '//Report file Creation
CLOSEFILE:
    If err.Number = 55 Then
        Close #1
        Open DirDstn & "\PLU.txt" For Output As #1 '//Report file Creation
    End If
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "Select * From ITEMMAST where LENGTH(PLU_CODE)>0", db, adOpenStatic, adLockReadOnly
    slcount = RSTCOMPANY.RecordCount
    Do Until RSTCOMPANY.EOF
        Print #1, IIf(IsNull(RSTCOMPANY!PLU_CODE), "", RSTCOMPANY!PLU_CODE) & "," & IIf(IsNull(RSTCOMPANY!ITEM_CODE), "", RSTCOMPANY!ITEM_CODE) & "," & IIf(IsNull(RSTCOMPANY!ITEM_NAME), "", RSTCOMPANY!ITEM_NAME) & "," & IIf(IsNull(RSTCOMPANY!PACK_TYPE) Or RSTCOMPANY!PACK_TYPE = "", "Kg", RSTCOMPANY!PACK_TYPE) & "," & IIf(IsNull(RSTCOMPANY!P_RETAIL), 0, RSTCOMPANY!P_RETAIL)
        RSTCOMPANY.MoveNext
    Loop
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    

    Print #1, Chr(13)
    
    'Print #1, Chr(27) & Chr(80)
    Close #1 '//Closing the file
    MsgBox slcount & " Items Updated", , "EzBiz"
    Exit Sub

eRRhAND:
    Screen.MousePointer = vbNormal
     MsgBox err.Description

End Sub

Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub DrvDest_Change()
    On Error GoTo eRRhAND
    DirDstn.Path = DrvDest
    Exit Sub
eRRhAND:
    If err.Number = 68 Then
        DrvDest = "C:\"
        DirDstn.Path = "C:\"
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub Form_Load()
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    On Error GoTo eRRhAND
    If FileExists(App.Path & "\PLUPATH") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\PLUPATH")  'Reading from the file
        DirDstn.Path = ObjFile.ReadLine
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    cetre Me
    Exit Sub
eRRhAND:
    If err.Number = 68 Then
        DrvDest = "C:\"
        DirDstn.Path = "C:\"
    Else
        MsgBox err.Description, , "EzBiz"
    End If
End Sub

