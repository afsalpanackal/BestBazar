VERSION 5.00
Begin VB.Form frmCalculator1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculator"
   ClientHeight    =   4710
   ClientLeft      =   7935
   ClientTop       =   3270
   ClientWidth     =   5130
   ForeColor       =   &H00000000&
   Icon            =   "frmCalculator1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4710
   ScaleWidth      =   5130
   Begin VB.CommandButton cmdCalc 
      Caption         =   "="
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   22
      Left            =   4020
      TabIndex        =   0
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   21
      Left            =   3060
      TabIndex        =   23
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   20
      Left            =   2100
      TabIndex        =   22
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   19
      Left            =   1140
      TabIndex        =   21
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   18
      Left            =   180
      TabIndex        =   20
      Top             =   3765
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1/x"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   17
      Left            =   4020
      TabIndex        =   19
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   16
      Left            =   3060
      TabIndex        =   18
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   15
      Left            =   2100
      TabIndex        =   17
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   14
      Left            =   1140
      TabIndex        =   16
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   13
      Left            =   180
      TabIndex        =   15
      Top             =   2985
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   12
      Left            =   4020
      TabIndex        =   14
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   11
      Left            =   3060
      TabIndex        =   13
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   10
      Left            =   2100
      TabIndex        =   12
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   9
      Left            =   1140
      TabIndex        =   11
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   8
      Left            =   180
      TabIndex        =   10
      Top             =   2220
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   7
      Left            =   4020
      TabIndex        =   9
      Top             =   1485
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   6
      Left            =   3060
      TabIndex        =   8
      Top             =   1485
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   5
      Left            =   2100
      TabIndex        =   7
      Top             =   1485
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   4
      Left            =   1140
      TabIndex        =   6
      Top             =   1485
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   700
      Index           =   3
      Left            =   180
      TabIndex        =   5
      Top             =   1485
      Width           =   915
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   2
      Left            =   3480
      TabIndex        =   4
      Top             =   885
      Width           =   1455
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "CE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   1860
      TabIndex        =   3
      Top             =   885
      Width           =   1515
   End
   Begin VB.CommandButton cmdCalc 
      Caption         =   "Backspace"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   885
      Width           =   1515
   End
   Begin VB.Label lblDisplay 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0."
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   750
      Left            =   165
      TabIndex        =   1
      Top             =   60
      Width           =   4785
   End
End
Attribute VB_Name = "frmCalculator1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mdblResult           As Double
Private mdblSavedNumber      As Double
Private mstrDot              As String
Private mstrOp               As String
Private mstrDisplay          As String
Private mblnDecEntered       As Boolean
Private mblnOpPending        As Boolean
Private mblnNewEquals        As Boolean
Private mblnEqualsPressed    As Boolean
Private mintCurrKeyIndex     As Integer
Private m_pass               As String

Private Sub Form_Load()

    Top = 0 '(Screen.Height - Height) / 2
    Left = (Screen.Width - Width) - 1200 '/ 2

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intIndex    As Integer
    
    Select Case KeyCode
        Case vbKeyBack:             intIndex = 0
        Case vbKeyDelete:           intIndex = 1
        Case vbKeyEscape:           intIndex = 2
        Case vbKey0, vbKeyNumpad0:  intIndex = 18
        Case vbKey1, vbKeyNumpad1:  intIndex = 13
        Case vbKey2, vbKeyNumpad2:  intIndex = 14
        Case vbKey3, vbKeyNumpad3:  intIndex = 15
        Case vbKey4, vbKeyNumpad4:  intIndex = 8
        Case vbKey5, vbKeyNumpad5:  intIndex = 9
        Case vbKey6, vbKeyNumpad6:  intIndex = 10
        Case vbKey7, vbKeyNumpad7:  intIndex = 3
        Case vbKey8, vbKeyNumpad8:  intIndex = 4
        Case vbKey9, vbKeyNumpad9:  intIndex = 5
        Case vbKeyDecimal, 190:     intIndex = 20
        Case vbKeyAdd:              intIndex = 21
        Case vbKeySubtract:         intIndex = 16
        Case vbKeyMultiply:         intIndex = 11
        Case vbKeyDivide:           intIndex = 6
        Case vbKeyReturn:           intIndex = 22
        Case Else:                  Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex
    
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim intIndex    As Integer
    
    Select Case Chr$(KeyAscii)
        Case "S", "s":  intIndex = 7
        Case "P", "p":  intIndex = 12
        Case "R", "r":  intIndex = 17
        Case "X", "x":  intIndex = 11
        Case "=":       intIndex = 22
        Case Else:      Exit Sub
    End Select
    
    cmdCalc(intIndex).SetFocus
    cmdCalc_Click intIndex

End Sub

Private Sub cmdCalc_Click(Index As Integer)
    
    Dim strPressedKey   As String
    
    mintCurrKeyIndex = Index
    
    If mstrDisplay = "ERROR" Then
        mstrDisplay = ""
    End If
    
    strPressedKey = cmdCalc(Index).Caption
    
    Select Case strPressedKey
        Case "0", "1", "2", "3", "4", "5", "6", _
             "7", "8", "9"
            m_pass = m_pass & strPressedKey
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            mstrDisplay = mstrDisplay & strPressedKey
        Case "."
            If mblnOpPending Then
                mstrDisplay = ""
                mblnOpPending = False
            End If
            If mblnEqualsPressed Then
                mstrDisplay = ""
                mblnEqualsPressed = False
            End If
            If InStr(mstrDisplay, ".") > 0 Then
                Beep
            Else
                mstrDisplay = mstrDisplay & strPressedKey
            End If
        Case "X"
            'If m_pass = "1548" Then 'MANGANTHANAM
            'If m_pass = "786" Then ' CEECEE
            'If m_pass = "9256" Then ' SB
            'If m_pass = "9847" Then '3 sTAR
            'If m_pass = "2070" Then 'Ge0
            'If m_pass = "4720" Then 'Common
'            If m_pass = "8984" Then 'Sujith
'                frmLogin.Show
'
'                Unload Me
'                Exit Sub
'            End If
            mdblResult = Val(mstrDisplay)
            mstrOp = strPressedKey
            mblnOpPending = True
            mblnDecEntered = False
            mblnNewEquals = True
        Case "+", "-", "/"
            mdblResult = Val(mstrDisplay)
            mstrOp = strPressedKey
            mblnOpPending = True
            mblnDecEntered = False
            mblnNewEquals = True
        Case "%"
            mdblSavedNumber = (Val(mstrDisplay) / 100) * mdblResult
            mstrDisplay = Format$(mdblSavedNumber)
        Case "="
            If mblnNewEquals Then
                mdblSavedNumber = Val(mstrDisplay)
                mblnNewEquals = False
            End If
            Select Case mstrOp
                Case "+"
                    mdblResult = mdblResult + mdblSavedNumber
                Case "-"
                    mdblResult = mdblResult - mdblSavedNumber
                Case "X"
                    mdblResult = mdblResult * mdblSavedNumber
                Case "/"
                    If mdblSavedNumber = 0 Then
                        mstrDisplay = "ERROR"
                    Else
                        mdblResult = mdblResult / mdblSavedNumber
                    End If
                Case Else
                    mdblResult = Val(mstrDisplay)
            End Select
            If mstrDisplay <> "ERROR" Then
                mstrDisplay = Format$(mdblResult)
            End If
            mblnEqualsPressed = True
        Case "+/-"
            If mstrDisplay <> "" Then
                If Left$(mstrDisplay, 1) = "-" Then
                    mstrDisplay = Right$(mstrDisplay, 2)
                Else
                    mstrDisplay = "-" & mstrDisplay
                End If
            End If
        Case "Backspace"
            If Val(mstrDisplay) <> 0 Then
                mstrDisplay = Left$(mstrDisplay, Len(mstrDisplay) - 1)
                mdblResult = Val(mstrDisplay)
            End If
            m_pass = ""
        Case "CE"
            mstrDisplay = ""
            m_pass = ""
        Case "C"
            mstrDisplay = ""
            mdblResult = 0
            mdblSavedNumber = 0
            m_pass = ""
        Case "1/x"
            If Val(mstrDisplay) = 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = 1 / mdblResult
                mstrDisplay = Format$(mdblResult)
            End If
        Case "sqrt"
            If Val(mstrDisplay) < 0 Then
                mstrDisplay = "ERROR"
            Else
                mdblResult = Val(mstrDisplay)
                mdblResult = Sqr(mdblResult)
                mstrDisplay = Format$(mdblResult)
            End If
    End Select
        
    If mstrDisplay = "" Then
        lblDisplay = "0."
    Else
        mstrDot = IIf(InStr(mstrDisplay, ".") > 0, "", ".")
        lblDisplay = mstrDisplay & mstrDot
        If Left$(lblDisplay, 1) = "0" Then
            lblDisplay = Mid$(lblDisplay, 2)
        End If
    End If
    
    If lblDisplay = "." Then lblDisplay = "0."
    cmdCalc(22).SetFocus
End Sub
