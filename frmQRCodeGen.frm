VERSION 5.00
Begin VB.Form frmQRCodeGen 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ultimate - QrCode Genarator Vb6 Code @ ARRATech Software Solution Ltd."
   ClientHeight    =   7125
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   7575
   Begin VB.CommandButton cmdSentMail 
      Caption         =   "SEnd Mail"
      Height          =   615
      Left            =   4080
      TabIndex        =   8
      Top             =   6090
      Width           =   3105
   End
   Begin VB.CommandButton cmdDirectPrint 
      Caption         =   "Direct Print"
      Height          =   855
      Left            =   4080
      TabIndex        =   7
      Top             =   4980
      Width           =   3075
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Height          =   585
      Left            =   4020
      TabIndex        =   6
      Top             =   4080
      Width           =   3105
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate QR"
      Height          =   585
      Left            =   3990
      TabIndex        =   5
      Top             =   3240
      Width           =   3285
   End
   Begin VB.TextBox txtNewQrCode 
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   3570
      Width           =   3345
   End
   Begin VB.TextBox txtLocation 
      Height          =   405
      Left            =   1000
      TabIndex        =   3
      Top             =   210
      Width           =   2895
   End
   Begin VB.ComboBox cboApplication 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   5580
      Width           =   3165
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00808000&
      Height          =   735
      Left            =   3960
      Top             =   2460
      Width           =   3315
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      Height          =   735
      Left            =   3960
      Stretch         =   -1  'True
      Top             =   2460
      Width           =   3315
   End
   Begin VB.Image Image1 
      Height          =   3045
      Left            =   270
      Picture         =   "frmQRCodeGen.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3300
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Id Number / Name etc."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   345
      Left            =   240
      TabIndex        =   2
      Top             =   3270
      Width           =   2805
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "File Format"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   465
      Left            =   270
      TabIndex        =   1
      Top             =   5010
      Width           =   2265
   End
   Begin VB.Image ImgQrCode 
      Appearance      =   0  'Flat
      Height          =   2265
      Left            =   3990
      Stretch         =   -1  'True
      Top             =   150
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808000&
      Height          =   2295
      Left            =   3960
      Top             =   120
      Width           =   3315
   End
End
Attribute VB_Name = "frmQRCodeGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim etsCrxRep As CRAXDRT.Application
Dim mcReport As CRAXDRT.Report
Public ImgLoc As String
Private Sub cmdSentMail_Click()
    frrmCompose.Show 1
End Sub

Private Sub Form_Load()

    cboApplication.AddItem "Acrobat Format PDF"
    cboApplication.AddItem "MS Word"
    cboApplication.AddItem "MS Excel"
    cboApplication.AddItem "Crystal Reports"
    cboApplication.AddItem "Vb6 Data Report"
    cboApplication.AddItem "Images"
    cboApplication.ListIndex = 4
End Sub
Private Sub cmdGenerate_click()
    Dim Con_char As String
   
    Dim UniSerial   As String
       
        UniSerial = Hour(Now) & Minute(Now) & Second(Now)
        ImgLoc = App.Path & "\QRCodeTemp\" & UniSerial & txtNewQrCode.text & ".jpg"
         
        Con_char = txtNewQrCode.text
        Call GenerateQrCode(ImgLoc, Con_char)
        DoEvents
        ImgQrCode.Picture = LoadPicture(ImgLoc)
        
        DoEvents
        
    
    
End Sub
Public Sub GenerateQrCode(ByRef pLoc As String, ByRef QrSerial As String)
    GenerateBMP StrPtr(pLoc), StrPtr(QrSerial), 3, 5, QualityLow
End Sub
Private Sub cmdDirectPrint_Click()
        Dim crAckSP         As New rptCusSPAcknowledgement
        
       
       
         With crAckSP

            .PaperSize = crPaperUser
            .PaperOrientation = crPortrait
            
            .ParameterFields(1).ClearCurrentValueAndRange
            .ParameterFields(1).AddCurrentValue txtNewQrCode.text
            
            .PrintOut False, 1
      
        End With
          
       
       
End Sub
Private Sub cmdPreview_Click()
    Dim crAckSP As New rptCusSPAcknowledgement
     
    Call ExportReportToFormat(crAckSP, txtNewQrCode.text, txtNewQrCode)
End Sub
Private Sub ExportReportToFormat(ReportObject As CRAXDRT.Report, ByVal FileName As String, ByVal ReportTitle As String)
   
    If cboApplication.text = "MS Excel" Then
        Call InsertPicture(txtLocation.text, App.Path & "\FileTemp\" & FileName & ".xls")
        ShellExecute 0, "open", App.Path & "\FileTemp\" & FileName & ".xls", "", "", vbNormalFocus
        Exit Sub
    End If
    If cboApplication.text = "Images" Then
        ShellExecute 0, "open", txtLocation.text, "", "", vbNormalFocus
        Exit Sub
    End If
    If cboApplication.text = "MS Word" Then
        Call InsertPictureWord(txtLocation.text, App.Path & "\FileTemp\" & FileName & ".doc")
        Exit Sub
    End If
    
    If cboApplication.text = "Vb6 Data Report" Then
        DataReport1.Show 1
        Exit Sub
    End If
    
    
      
    Dim crAckSP As New rptCusSPAcknowledgement
    
    Dim objExportOptions As CRAXDRT.ExportOptions
    Dim Extn As String
    
    
    ReportObject.ReportTitle = ReportTitle
    
    With ReportObject
        .EnableParameterPrompting = False
        .MorePrintEngineErrorMessages = True
    End With
    
    Set objExportOptions = ReportObject.ExportOptions
    
    With objExportOptions
        .DestinationType = crEDTDiskFile
        '.DestinationType = crEDTApplication
        
        Select Case cboApplication.text

            Case "Acrobat Format PDF"
                .FormatType = crEFTPortableDocFormat
                Extn = ".pdf"
            Case "Crystal Reports"
                 PRNmc.Show 1
                 Exit Sub
        End Select
             .DiskFileName = App.Path & "\FileTemp\" & FileName & Extn
             .PDFExportAllPages = True
            
    End With
 
    ReportObject.Export False 'True
    ShellExecute 0, "open", App.Path & "\FileTemp\" & FileName & Extn, "", "", vbNormalFocus

End Sub
Sub InsertPicture(Picpath As String, FileDestination As String)
        Dim AppExcel As Excel.Application
        Dim AppBook As Excel.Workbook
        Dim AppSheet As Excel.Worksheet
            
        Dim myPic As Object
       
            
        'Start a new workbook in Excel
        Set AppExcel = CreateObject("Excel.Application")
        Set AppBook = AppExcel.Workbooks.Add
              
        'Add data to cells of the first worksheet in the new workbook
        Set AppSheet = AppBook.Worksheets(1)
        AppExcel.Range("E1").Value = txtNewQrCode.text
      
            
            Set myPic = AppSheet.Shapes.AddPicture(Picpath, False, True, 0, 5, -1, -1)
            
            With myPic
                
                .Width = 75
                .Height = 100
                .Top = AppExcel.Cells(2, 8).Top 'according to variables from correct answer
                .Left = AppExcel.Cells(5, 5).Left
                .LockAspectRatio = msoFalse
                
            End With
        
        'Save the Workbook and Quit Excel
        AppBook.SaveAs FileDestination
        AppExcel.Quit
        
End Sub
Sub InsertPictureWord(Picpath As String, FileDestination As String)
   
    'Dim r%, c%, pct As PictureBox
    Dim objWordApp As Word.Application
    Dim objWordDoc As Word.Document
    Dim objRange As Word.Range
    Dim objTable As Word.Table
    
    Set objWordApp = New Word.Application
    Set objWordDoc = objWordApp.Documents.Add
    
    '-----------------------------------------------------------------
    'insert image from file
    '-----------------------------------------------------------------
    
    'Set pct = Me.Controls.Add("VB.Picturebox", "pctTemp")
    'pct.Picture = LoadPicture(Picpath) '<-- specify correct file path and name here
    
    Clipboard.Clear
    Clipboard.SetData Me.ImgQrCode, vbCFBitmap
    objWordApp.Selection.Paste
    
    objWordApp.Selection.TypeText vbNewLine & vbNewLine
       
    objWordApp.Visible = True
    
    Set objWordDoc = Nothing
    Set objWordApp = Nothing
   

End Sub

'===========================================================================
'Additional References
'Private Sub cmdPrtProductionSheet_Click()
'
'    On Error GoTo cmbPrtProductionSheet_Error:
'
'    Dim objWORD As New Word.Application
'    Dim objDoc As New Word.Document
'
''---------------------------------- Output Header -----------------------------------
'
'    Set objDoc = objWORD.Documents.Open("C:\PPSheet.doc")
'
'    objDoc.FormFields.Item("clientcode").Range = strClientCode
'    objDoc.FormFields.Item("factorycode").Range = txtFactoryCode
'    objDoc.FormFields.Item("PPNO").Range = strPPNO
'    objDoc.FormFields.Item("PDate").Range = strIDate
'    objDoc.FormFields.Item("EDate").Range = strEDate
'    objDoc.FormFields.Item("ContractNo").Range = strSONO
'    objDoc.FormFields.Item("StyleCode").Range = txtPPStyleCode
'    objDoc.FormFields.Item("BulkQty").Range = txtPPBulkQty
'    objDoc.FormFields.Item("Quota").Range = txtPPQuota
'    objDoc.FormFields.Item("description").Range = txtPPStyleName
'
'------------------------------ Output Image File -------------------------------------
'
'    If iFile = vbNullString Then GoTo cmbPrtProductionSheet_DataGrid:
'
'    Dim r%, c%
'    Dim pct As PictureBox
'
'    Set pct = Me.Controls.Add("VB.Picturebox", "pctTemp")
'
'    pct.Picture = LoadPicture(iFile)
'
'    Clipboard.Clear
'    Clipboard.SetData pct.Image, vbCFBitmap
'
'    objWORD.Selection.Paste
'    objWORD.Selection.TypeText vbNewLine & vbNewLine
'
'    Me.Controls.Remove "pctTemp"
'
''------------------------------- Output DataGrid -------------------------------------
'
'cmbPrtProductionSheet_DataGrid:
'
'    Dim objRange As Word.Range
'    Dim objTable As Word.Table
'    Dim RS As ADODB.Recordset
'
'    Set RS = New ADODB.Recordset
'
'    RS.CursorLocation = adUseClient
'
'    Set RS = Adodc1.Recordset
'
'    If RS.EOF Then GoTo cmbPrtProductionSheet_Exit:
'
'    r = RS.RecordCount
'    c = RS.Fields.Count
'
'    For i = 36 To 6 Step -1
'        If RS.Fields(i).Value <> vbNullString Then
'            j = i - 4
'            i = 6
'        End If
'    Next i
'
'    Set objRange = objWORD.Selection.Range
'    Set objTable = objDoc.Tables.Add(objRange, r, j)
'
'    With RS
'        r = 1
'        Do While Not .EOF
'            For c = 1 To j
'                objTable.Cell(r, c).Range.text = vbNullString & .Fields(c + 4).Value
'            Next c
'            .MoveNext
'            r = r + 1
'        Loop
'    End With
'
'cmbPrtProductionSheet_Exit:
'    On Error Resume Next
'    objDoc.SaveAs ("C:\PPSheet_" & strPPNO & ".doc")
'    objDoc.Close
'    Set RS = Nothing
'    Set pct = Nothing
'    Set objDoc = Nothing
'    Set objWORD = Nothing
'    Exit Sub
'
'cmbPrtProductionSheet_Error:
'    Resume cmbPrtProductionSheet_Exit:
'
'End Sub

