VERSION 5.00
Object = "{F62B9FA4-455F-4FE3-8A2D-205E4F0BCAFB}#11.5#0"; "CRViewer.dll"
Begin VB.Form frmreport 
   Caption         =   "Report"
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12345
   Icon            =   "frmreport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   9360
   ScaleWidth      =   12345
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdmail 
      Caption         =   "&Email"
      Height          =   405
      Left            =   9780
      TabIndex        =   2
      Top             =   15
      Width           =   975
   End
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   405
      Left            =   8745
      TabIndex        =   1
      Top             =   15
      Width           =   975
   End
   Begin VB.CommandButton CmdPrint 
      Caption         =   "&Print"
      Height          =   405
      Left            =   7695
      TabIndex        =   0
      Top             =   15
      Width           =   975
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer91 
      Height          =   2640
      Left            =   1695
      TabIndex        =   3
      Top             =   1875
      Width           =   5535
      _cx             =   9763
      _cy             =   4657
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   2057
      EnableInteractiveParameterPrompting=   0   'False
   End
End
Attribute VB_Name = "frmreport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''First make sure to include the references:
''Crystal Report Viewer Control 9
''Crystal Report 9 ActiveX Designer Runtime Library
''
''You need a form where you have added the control Crystal Report Viewer Control 9. I have called it frmReport in my module.
''The control will take up the whole form.
''Add this code to that form:
''
''Code:

Option Explicit

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub CmdPrint_Click()
    CRViewer91.PrintReport
    'Unload Me
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    CmdPrint.SetFocus
End Sub

'=========================================================================================
Private Sub Form_Load()
    CRViewer91.EnableExportButton = True
    CRViewer91.EnableGroupTree = False
    CRViewer91.DisplayTabs = False
    CRViewer91.EnableCloseButton = False
    CRViewer91.EnableProgressControl = True
End Sub 'Form_Load()
'=========================================================================================
Private Sub Form_Resize()
    CRViewer91.Top = 0
    CRViewer91.Left = 0
    CRViewer91.Height = ScaleHeight
    'CRViewer91.WhatsThisHelpID = ScaleWidth
    CRViewer91.Width = ScaleWidth
End Sub 'Form_Resize()

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.Enabled = True
End Sub

Public Sub ViewReport()
    'View the Report
    CRViewer91.ViewReport
End Sub

