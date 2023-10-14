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
   ScaleHeight     =   12375
   ScaleWidth      =   22800
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      BackColor       =   &H00CAF1F9&
      Caption         =   "Save as &Excel"
      Default         =   -1  'True
      Height          =   405
      Left            =   12495
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   15
      Width           =   1290
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00ECD1F5&
      Caption         =   "&Save as PDF"
      Height          =   405
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   15
      Width           =   1290
   End
   Begin VB.CommandButton CmdPrint2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Print &2 Copies"
      Height          =   405
      Left            =   8880
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   15
      Width           =   1230
   End
   Begin VB.CommandButton CmdClose 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Close"
      Height          =   405
      Left            =   10140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   15
      Width           =   975
   End
   Begin VB.CommandButton CmdPrint 
      BackColor       =   &H00FFC0C0&
      Caption         =   "&Print"
      Height          =   405
      Left            =   7860
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   15
      Width           =   975
   End
   Begin CrystalActiveXReportViewerLib11_5Ctl.CrystalActiveXReportViewer CRViewer91 
      Height          =   2640
      Left            =   1695
      TabIndex        =   2
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
    'CRViewer91.PrintReport
    'Unload Me
    Report.PrintOut False
    CmdClose.SetFocus
End Sub

Private Sub CmdPrint2_Click()
    Report.PrintOut False, 2
    CmdClose.SetFocus
End Sub

Private Sub Command1_Click()
    Call saveas(0)
End Sub

Private Sub Command2_Click()
    Call saveas(1)
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
    Set Report = Nothing
    MDIMAIN.Enabled = True
End Sub

Public Sub ViewReport()
    'View the Report
    CRViewer91.ViewReport
End Sub

Private Function saveas(exptype As Integer)
    
    Dim oxopt As CRAXDRT.ExportOptions
    Dim filepath As String
    
    On Error GoTo ERRHAND
    Set oxopt = Report.ExportOptions
    With oxopt
        .DestinationType = crEDTDiskFile
        Select Case exptype
            Case 1
                .DiskFileName = App.Path & "\Report.xls"
                .FormatType = crEFTExcel97
                filepath = App.Path & "\Report.xls"
            Case Else
                .DiskFileName = App.Path & "\Report.PDF"
                .FormatType = crEFTPortableDocFormat
                filepath = App.Path & "\Report.PDF"
        End Select
        
    End With
    Report.Export False
    
    On Error Resume Next
    If Trim(filepath) <> "" Then Debug.Print ShellExecute(hwnd, "open", Trim(filepath), vbNullString, vbNullString, 1)
    err.Clear
    Exit Function
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Function
