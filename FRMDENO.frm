VERSION 5.00
Begin VB.Form FrmDenom 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Denomination"
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5385
   Icon            =   "FRMDENO.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtDiff 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1320
      Locked          =   -1  'True
      TabIndex        =   59
      Top             =   7140
      Width           =   1770
   End
   Begin VB.CommandButton CmdReset 
      Caption         =   "&Reset"
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
      Left            =   75
      TabIndex        =   58
      Top             =   7155
      Width           =   1215
   End
   Begin VB.TextBox TxtCAmount 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   915
      TabIndex        =   11
      Top             =   6540
      Width           =   2190
   End
   Begin VB.TextBox TxtResult 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   57
      Top             =   6540
      Width           =   1785
   End
   Begin VB.TextBox LblHalf 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   5820
      Width           =   1785
   End
   Begin VB.TextBox Txthalf 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   10
      Top             =   5820
      Width           =   1005
   End
   Begin VB.TextBox Lbl1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   55
      Top             =   5250
      Width           =   1785
   End
   Begin VB.TextBox txt1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   9
      Top             =   5250
      Width           =   1005
   End
   Begin VB.TextBox Lbl2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   54
      Top             =   4680
      Width           =   1785
   End
   Begin VB.TextBox Txt2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   8
      Top             =   4680
      Width           =   1005
   End
   Begin VB.TextBox Lbl5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   53
      Top             =   4110
      Width           =   1785
   End
   Begin VB.TextBox Txt5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   7
      Top             =   4110
      Width           =   1005
   End
   Begin VB.TextBox Lbl10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   52
      Top             =   3540
      Width           =   1785
   End
   Begin VB.TextBox Txt10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   6
      Top             =   3540
      Width           =   1005
   End
   Begin VB.TextBox Lbl20 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   51
      Top             =   2970
      Width           =   1785
   End
   Begin VB.TextBox Txt20 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   5
      Top             =   2970
      Width           =   1005
   End
   Begin VB.TextBox Lbl50 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   2400
      Width           =   1785
   End
   Begin VB.TextBox Txt50 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   4
      Top             =   2400
      Width           =   1005
   End
   Begin VB.TextBox Lbl100 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   49
      Top             =   1815
      Width           =   1785
   End
   Begin VB.TextBox Txt100 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   3
      Top             =   1815
      Width           =   1005
   End
   Begin VB.TextBox Lbl200 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   48
      Top             =   1245
      Width           =   1785
   End
   Begin VB.TextBox Txt200 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   2
      Top             =   1245
      Width           =   1005
   End
   Begin VB.TextBox Lbl500 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   47
      Top             =   675
      Width           =   1785
   End
   Begin VB.TextBox Txt500 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   1
      Top             =   675
      Width           =   1005
   End
   Begin VB.TextBox Lbl2000 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   3420
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   1785
   End
   Begin VB.TextBox Txt2000 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1965
      TabIndex        =   0
      Top             =   120
      Width           =   1005
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
      Height          =   540
      Left            =   4035
      TabIndex        =   12
      Top             =   7155
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   32
      Left            =   3015
      TabIndex        =   46
      Top             =   4125
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   31
      Left            =   3015
      TabIndex        =   45
      Top             =   2985
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   30
      Left            =   3015
      TabIndex        =   44
      Top             =   5850
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   29
      Left            =   3015
      TabIndex        =   43
      Top             =   5295
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   28
      Left            =   3015
      TabIndex        =   42
      Top             =   4695
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   27
      Left            =   3015
      TabIndex        =   41
      Top             =   3540
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   26
      Left            =   3015
      TabIndex        =   40
      Top             =   2415
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   25
      Left            =   3015
      TabIndex        =   39
      Top             =   1830
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   24
      Left            =   3015
      TabIndex        =   38
      Top             =   1260
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   23
      Left            =   3015
      TabIndex        =   37
      Top             =   690
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   22
      Left            =   3015
      TabIndex        =   36
      Top             =   135
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   21
      Left            =   1500
      TabIndex        =   35
      Top             =   4125
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   20
      Left            =   1500
      TabIndex        =   34
      Top             =   2985
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   19
      Left            =   1500
      TabIndex        =   33
      Top             =   5850
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   18
      Left            =   1500
      TabIndex        =   32
      Top             =   5295
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   17
      Left            =   1500
      TabIndex        =   31
      Top             =   4695
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   16
      Left            =   1500
      TabIndex        =   30
      Top             =   3540
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   15
      Left            =   1500
      TabIndex        =   29
      Top             =   2415
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   14
      Left            =   1500
      TabIndex        =   28
      Top             =   1830
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   13
      Left            =   1500
      TabIndex        =   27
      Top             =   1260
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   12
      Left            =   1500
      TabIndex        =   26
      Top             =   690
      Width           =   250
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Index           =   11
      Left            =   1500
      TabIndex        =   25
      Top             =   135
      Width           =   250
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50 (P)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   10
      Left            =   60
      TabIndex        =   24
      Top             =   5910
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   9
      Left            =   210
      TabIndex        =   23
      Top             =   5310
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   8
      Left            =   210
      TabIndex        =   22
      Top             =   4740
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   7
      Left            =   240
      TabIndex        =   21
      Top             =   4230
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   6
      Left            =   240
      TabIndex        =   20
      Top             =   3630
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   3060
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "50"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   2460
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   3
      Left            =   255
      TabIndex        =   17
      Top             =   1920
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "200"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   2
      Left            =   255
      TabIndex        =   16
      Top             =   1320
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   1
      Left            =   255
      TabIndex        =   15
      Top             =   750
      Width           =   1100
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Index           =   0
      Left            =   255
      TabIndex        =   13
      Top             =   150
      Width           =   1100
   End
End
Attribute VB_Name = "FrmDenom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub CmdReset_Click()
    TxtResult.Visible = False
    Lbl2000.Text = ""
    Lbl500.Text = ""
    Lbl200.Text = ""
    Lbl100.Text = ""
    Lbl50.Text = ""
    Lbl20.Text = ""
    Lbl10.Text = ""
    Lbl5.Text = ""
    Lbl2.Text = ""
    Lbl1.Text = ""
    LblHalf.Text = ""
    
    Txt2000.Text = ""
    Txt500.Text = ""
    Txt200.Text = ""
    Txt100.Text = ""
    Txt50.Text = ""
    Txt20.Text = ""
    Txt10.Text = ""
    Txt5.Text = ""
    Txt2.Text = ""
    txt1.Text = ""
    Txthalf.Text = ""
    
    TxtDiff.Text = ""
    TxtResult.Text = ""
    TxtResult.Visible = True
End Sub

Private Sub Form_Load()
    Me.Top = 100
    Me.Left = 10000
End Sub

Private Sub txt1_Change()
    If TxtResult.Visible = True Then
        Lbl1.Text = Val(txt1.Text) * 1
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub txt1_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txthalf.SetFocus
        Case vbKeyEscape
            Txt2.SetFocus
    End Select
End Sub

Private Sub txt1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt10_Change()
    If TxtResult.Visible = True Then
        Lbl10.Text = Val(Txt10.Text) * 10
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt10_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt5.SetFocus
        Case vbKeyEscape
            Txt20.SetFocus
    End Select
End Sub

Private Sub Txt10_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt100_Change()
    If TxtResult.Visible = True Then
        Lbl100.Text = Val(Txt100.Text) * 100
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt100_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt50.SetFocus
        Case vbKeyEscape
            Txt200.SetFocus
    End Select
End Sub

Private Sub Txt100_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt2_Change()
    If TxtResult.Visible = True Then
        Lbl2.Text = Val(Txt2.Text) * 2
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt2_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            txt1.SetFocus
        Case vbKeyEscape
            Txt5.SetFocus
    End Select
End Sub

Private Sub Txt2_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt20_Change()
    If TxtResult.Visible = True Then
        Lbl20.Text = Val(Txt20.Text) * 20
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt20_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt10.SetFocus
        Case vbKeyEscape
            Txt50.SetFocus
    End Select
End Sub

Private Sub Txt20_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt200_Change()
    If TxtResult.Visible = True Then
        Lbl200.Text = Val(Txt200.Text) * 200
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt200_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt100.SetFocus
        Case vbKeyEscape
            Txt500.SetFocus
    End Select
End Sub

Private Sub Txt200_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt2000_Change()
    If TxtResult.Visible = True Then
        Lbl2000.Text = Val(Txt2000.Text) * 2000
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt2000_GotFocus()
    Txt2000.SelStart = 0
    Txt2000.SelLength = Len(Txt2000.Text)
End Sub

Private Sub Txt500_GotFocus()
    Txt500.SelStart = 0
    Txt500.SelLength = Len(Txt500.Text)
End Sub

Private Sub Txt200_GotFocus()
    Txt200.SelStart = 0
    Txt200.SelLength = Len(Txt200.Text)
End Sub

Private Sub Txt100_GotFocus()
    Txt100.SelStart = 0
    Txt100.SelLength = Len(Txt100.Text)
End Sub

Private Sub Txt50_GotFocus()
    Txt50.SelStart = 0
    Txt50.SelLength = Len(Txt50.Text)
End Sub

Private Sub Txt20_GotFocus()
    Txt20.SelStart = 0
    Txt20.SelLength = Len(Txt20.Text)
End Sub

Private Sub Txt10_GotFocus()
    Txt10.SelStart = 0
    Txt10.SelLength = Len(Txt10.Text)
End Sub

Private Sub Txt5_GotFocus()
    Txt5.SelStart = 0
    Txt5.SelLength = Len(Txt5.Text)
End Sub

Private Sub Txt2_GotFocus()
    Txt2.SelStart = 0
    Txt2.SelLength = Len(Txt2.Text)
End Sub

Private Sub Txt1_GotFocus()
    txt1.SelStart = 0
    txt1.SelLength = Len(txt1.Text)
End Sub

Private Sub TxtHALF_GotFocus()
    Txthalf.SelStart = 0
    Txthalf.SelLength = Len(Txthalf.Text)
End Sub

Private Sub Txt2000_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt500.SetFocus
        Case vbKeyEscape
            
    End Select
End Sub

Private Sub Txt2000_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt5_Change()
    If TxtResult.Visible = True Then
        Lbl5.Text = Val(Txt5.Text) * 5
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt5_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt2.SetFocus
        Case vbKeyEscape
            Txt10.SetFocus
    End Select
End Sub

Private Sub Txt5_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt50_Change()
    If TxtResult.Visible = True Then
        Lbl50.Text = Val(Txt50.Text) * 50
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt50_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt20.SetFocus
        Case vbKeyEscape
            Txt100.SetFocus
    End Select
End Sub

Private Sub Txt50_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txt500_Change()
    If TxtResult.Visible = True Then
        Lbl500.Text = Val(Txt500.Text) * 500
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txt500_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            Txt200.SetFocus
        Case vbKeyEscape
            Txt2000.SetFocus
    End Select
End Sub

Private Sub Txt500_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub Txthalf_Change()
    If TxtResult.Visible = True Then
        LblHalf.Text = Val(Txthalf.Text) * 0.5
        TxtResult.Text = Val(Lbl2000.Text) + Val(Lbl500.Text) + Val(Lbl200.Text) + Val(Lbl100.Text) + Val(Lbl50.Text) + Val(Lbl20.Text) + Val(Lbl10.Text) + Val(Lbl5.Text) + Val(Lbl2.Text) + Val(Lbl1.Text) + Val(LblHalf.Text)
        TxtDiff.Text = Val(TxtCAmount.Text) - Val(TxtResult.Text)
    End If
End Sub

Private Sub Txthalf_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            
        Case vbKeyEscape
            txt1.SetFocus
    End Select
End Sub

Private Sub Txthalf_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub
