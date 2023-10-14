VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3750
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   5880
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":72FA
   ScaleHeight     =   3750
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtyear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   4
      TabIndex        =   2
      Top             =   1440
      Width           =   1170
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4530
      TabIndex        =   4
      Top             =   2670
      Width           =   1245
   End
   Begin VB.CommandButton cmdLogin 
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   3120
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2670
      Width           =   1275
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   825
      Width           =   3945
   End
   Begin VB.TextBox txtLogin 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      Left            =   1815
      MaxLength       =   20
      TabIndex        =   0
      Top             =   240
      Width           =   3945
   End
   Begin VB.Label LBLCOMP 
      BackStyle       =   0  'Transparent
      Caption         =   "COMPANY"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2055
      Width           =   1605
   End
   Begin MSForms.ComboBox CmbDB 
      Height          =   420
      Left            =   1830
      TabIndex        =   10
      Top             =   1995
      Width           =   3975
      ForeColor       =   0
      DisplayStyle    =   7
      Size            =   "7011;741"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "Tahoma"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "-"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   1470
      Width           =   210
   End
   Begin VB.Label lblfinyear 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   390
      Left            =   3255
      TabIndex        =   8
      Top             =   1425
      Width           =   1170
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Financial Year"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   495
      Left            =   135
      TabIndex        =   7
      Top             =   1470
      Width           =   1635
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   570
      Left            =   135
      TabIndex        =   6
      Top             =   900
      Width           =   1455
   End
   Begin VB.Label Login 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   450
      Left            =   135
      TabIndex        =   5
      Top             =   240
      Width           =   1845
   End
   Begin VB.Menu mnu 
      Caption         =   "Tools"
      Begin VB.Menu mnurestore 
         Caption         =   "Restore Database"
      End
      Begin VB.Menu mnufix 
         Caption         =   "Fix Table Error"
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs As New ADODB.Recordset
Private try As Integer

Private Sub CmdExit_Click()
    End
End Sub

Private Sub cmdLogin_Click()

'Set rs = New ADODB.Recordset
Dim sql As String
Dim MD5 As New clsMD5, Old_Password As String

On Error GoTo ERRHAND
If Len(txtyear.Text) <> 4 Then
    MsgBox "Please enter proper year", vbOKOnly, "LOGIN"
    txtyear.SetFocus
    Exit Sub
End If

If Val(txtyear.Text) < 2010 Or Val(txtyear.Text) > 2030 Then
    'MsgBox "Please enter proper year", vbOKOnly, "LOGIN"
    MsgBox "Unexpected error occured", vbOKOnly, "LOGIN"
    txtyear.SetFocus
    Exit Sub
End If

'If Month(Date) = 3 And Year(Date) >= 2021 And Day(Date) >= 15 Then
'    'db.Execute "delete from Users"
'    MsgBox "Annual Service Package expiers soon!! Please renew your Annual Service Package", vbOKOnly
'End If
'
'If (Year(Date) >= 2021 And Month(Date) >= 4) Then
'    MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly
'    'Exit Sub
'    'db.Execute "Delete from Users"
'End If

'If (Year(Date) >= 2020 And Month(Date) >= 9 And Day(Date) >= 12) Then
'    db.Execute "Delete from act_ky"
'End If

'If FileLen("D:\dbase\" & txtyear.Text) <> 0 Then Exit Sub

'If Not (Month(Date) = 8 Or Month(Date) = 9) Then
'    db.Execute "delete from Users"
'End If
'
'If FileLen("C:\windows\system32\sysfile.dll") <> 6 Then
'    db.Execute "delete from Users"
'End If


'If Len(Trim(txtLogin.Text)) < 2 Then GoTo errHand

'DBPwd = "DATA_RET"
'sql = "select * from USERS "
'rs.Open sql, db, adOpenKeyset, adLockPessimistic
'If rs.BOF And rs.EOF Then
'    MsgBox "Abnormal program termination. Crxdtl.dll missing", vbOKOnly, "EzBiz"
'    rs.Close
'    Set rs = Nothing
'    End
'End If
'rs.Close
'Set rs = Nothing
If CmbDB.Visible = True And CmbDB.Text <> "" And UCase(dbase1) <> "INV" & UCase(CmbDB.Text) Then
    db.Close
    Set db = Nothing
    dbase1 = "inv" & CmbDB.Text
    Dim strConn As String
    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase1 & ";User=root; Password=###%%database%%###ret; Option=2;"
    db.Open strConn
    db.CursorLocation = adUseClient
    
    strConnection = "Provider=MSDASQL;Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase1 & ";User=root; Password=###%%database%%###ret"
End If
Dim ObjFso
Dim StrFileName
Dim ObjFile
Dim printtype As String
    

Old_Password = UCase(MD5.DigestStrToHexStr(txtPassword.Text))
sql = "select * from USERS where USER_NAME='" & txtLogin.Text & "' and PASS_WORD='" & Old_Password & "'"
rs.Open sql, db, adOpenKeyset, adLockPessimistic
If rs.BOF And rs.EOF Then
    MsgBox "Incorrect login or password! Please try again...", vbOKOnly, "LOGIN...."
    rs.Close
    Set rs = Nothing
    txtPassword.SetFocus
    try = try + 1
    If try >= 5 Then
        MsgBox "You have exceed the limit", vbOKOnly, "Login"
        
        db.Close
        Set db = Nothing
        
'        dbprint.Close
'        Set dbprint = Nothing
        End
    End If
    
    Exit Sub
Else
    If Val(CALCODE) <> 0 Then
        If FileExists(App.Path & "\ico") Then MDIMAIN.Icon = LoadPicture(App.Path & "\ico")
    End If
    MDIMAIN.StatusBar.Panels(1).Text = "User: " & txtLogin.Text
    MDIMAIN.StatusBar.Panels(2).Text = "Last Login : " & IIf(IsNull(rs!LAST_LOGIN), "", rs!LAST_LOGIN)
    MDIMAIN.StatusBar.Panels(3).Text = "Last Logout : " & IIf(IsNull(rs!LAST_LOGOUT), "", rs!LAST_LOGOUT)
    MDIMAIN.StatusBar.Panels(4).Text = "Current Login : " & Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss")
    
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & txtyear.Text & "'", db, adOpenStatic, adLockOptimistic, adCmdText
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        MDIMAIN.StatusBar.Panels(5).Text = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
        customercode = IIf(IsNull(RSTCOMPANY!CUST_CODE), "", RSTCOMPANY!CUST_CODE)
        serveraddress = IIf(IsNull(RSTCOMPANY!SERVER_ADD), "", RSTCOMPANY!SERVER_ADD)
        REMOTEAPP = IIf(IsNull(RSTCOMPANY!CLOUD_TYPE), 0, RSTCOMPANY!CLOUD_TYPE)
        If RSTCOMPANY!REMOTE_FLAG = "Y" Then
            remoteflag = True
            MDIMAIN.mnusync.Visible = True
        Else
            remoteflag = False
            MDIMAIN.mnusync.Visible = False
        End If
        DUPCODE = IIf(IsNull(RSTCOMPANY!PCODE), "", RSTCOMPANY!PCODE)
        MDIMAIN.lblec.Caption = IIf(IsNull(RSTCOMPANY!EC), "", RSTCOMPANY!EC)
        MDIMAIN.StatusBar.Panels(8).Text = IIf(IsNull(RSTCOMPANY!DMP_FLAG), "N", RSTCOMPANY!DMP_FLAG)
        MDIMAIN.LBLSPACE.Caption = IIf(IsNull(RSTCOMPANY!LINE_SPACE), "N", RSTCOMPANY!LINE_SPACE)
        MDIMAIN.LBLDMPTHERMAL.Caption = IIf(IsNull(RSTCOMPANY!DMPTH_FLAG), "N", RSTCOMPANY!DMPTH_FLAG)
        MDIMAIN.StatusBar.Panels(9).Text = IIf(IsNull(RSTCOMPANY!DUP_FLAG), "N", RSTCOMPANY!DUP_FLAG)
        Roundflag = IIf(IsNull(RSTCOMPANY!ROUND_FLAG) Or RSTCOMPANY!ROUND_FLAG <> "N", True, False)
        BATCH_DISPLAY = IIf(IsNull(RSTCOMPANY!BATCH_FLAG) Or RSTCOMPANY!BATCH_FLAG <> "Y", False, True)
        MDIMAIN.StatusBar.Panels(10).Text = IIf(IsNull(RSTCOMPANY!Copy_8B), "1", RSTCOMPANY!Copy_8B)
        MDIMAIN.StatusBar.Panels(11).Text = IIf(IsNull(RSTCOMPANY!Copy_8), "1", RSTCOMPANY!Copy_8)
        MDIMAIN.StatusBar.Panels(12).Text = IIf(IsNull(RSTCOMPANY!Copy_8V), "1", RSTCOMPANY!Copy_8V)
        MDIMAIN.LBLTRCopy.Caption = IIf(IsNull(RSTCOMPANY!Copy_TR), "1", RSTCOMPANY!Copy_TR)
        MDIMAIN.StatusBar.Panels(13).Text = IIf(IsNull(RSTCOMPANY!PREVIEW_FLAG), "N", RSTCOMPANY!PREVIEW_FLAG)
        MDIMAIN.LBLTHPREVIEW.Caption = IIf(IsNull(RSTCOMPANY!PREVIEWTH_FLAG), "N", RSTCOMPANY!PREVIEWTH_FLAG)
        MDIMAIN.StatusBar.Panels(14).Text = IIf(IsNull(RSTCOMPANY!TAX_FLAG), "N", RSTCOMPANY!TAX_FLAG)
        MDIMAIN.StatusBar.Panels(15).Text = IIf(IsNull(RSTCOMPANY!CODE_FLAG), "N", RSTCOMPANY!CODE_FLAG)
        MDIMAIN.StatusBar.Panels(16).Text = IIf(IsNull(RSTCOMPANY!DISC_FLAG), "N", RSTCOMPANY!DISC_FLAG)
        MDIMAIN.StatusBar.Panels(6).Text = IIf(IsNull(RSTCOMPANY!BARCODE_FLAG), "N", RSTCOMPANY!BARCODE_FLAG)
        MDIMAIN.LBLTAXWARN.Caption = IIf(IsNull(RSTCOMPANY!TAXWRN_FLAG), "N", RSTCOMPANY!TAXWRN_FLAG)
        MDIMAIN.LBLAMC.Caption = IIf(IsNull(RSTCOMPANY!AMC_FLAG), "N", RSTCOMPANY!AMC_FLAG)
        hold_thermal = IIf(RSTCOMPANY!HOLD_THERMAL_FLAG = "Y", True, False)
        MDIMAIN.LBLGSTWRN.Caption = IIf(IsNull(RSTCOMPANY!GSTWRN_FLAG), "N", RSTCOMPANY!GSTWRN_FLAG)
        MDIMAIN.LBLITMWRN.Caption = IIf(IsNull(RSTCOMPANY!ITEM_WARN), "Y", RSTCOMPANY!ITEM_WARN)
        MDIMAIN.lblform62.Caption = IIf(IsNull(RSTCOMPANY!FORM_62), "N", RSTCOMPANY!FORM_62)
        MDIMAIN.lblgst.Caption = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
        MDIMAIN.lblnostock.Caption = IIf(IsNull(RSTCOMPANY!NSPT), "N", RSTCOMPANY!NSPT)
        MDIMAIN.lblprnall.Caption = IIf(IsNull(RSTCOMPANY!ALL_PRN), "N", RSTCOMPANY!ALL_PRN)
        MDIMAIN.lblRemoveUbill.Caption = IIf(IsNull(RSTCOMPANY!REMOVE_UBILL), "Y", RSTCOMPANY!REMOVE_UBILL)
        MRPDISC_FLAG = IIf(IsNull(RSTCOMPANY!MRP_DISC), "N", RSTCOMPANY!MRP_DISC)
        MDIMAIN.LBLSTATE.Caption = IIf(IsNull(RSTCOMPANY!SCODE) Or RSTCOMPANY!SCODE = 0 Or RSTCOMPANY!SCODE = "", "32", RSTCOMPANY!SCODE)
        MDIMAIN.LBLSTATENAME.Caption = IIf(IsNull(RSTCOMPANY!SNAME) Or RSTCOMPANY!SNAME = "", "KL", RSTCOMPANY!SNAME)
        If Trim(MDIMAIN.LBLSTATE.Caption) = "" Then MDIMAIN.LBLSTATE.Caption = "32"
        If Val(MDIMAIN.LBLSTATE.Caption) = 0 Then MDIMAIN.LBLSTATE.Caption = "32"
        If RSTCOMPANY!billtype_flag = "Y" Then
            bill_type_flag = True
        Else
            bill_type_flag = False
        End If
        
        If Len(MDIMAIN.LBLSTATE.Caption) <> 2 Then MDIMAIN.LBLSTATE.Caption = "32"
        If Val(MDIMAIN.LBLSTATE.Caption) = 32 Then MDIMAIN.LBLSTATENAME.Caption = "KL"
        MDIMAIN.lblExpEnable.Caption = IIf(IsNull(RSTCOMPANY!EXP_ENABLED), "N", RSTCOMPANY!EXP_ENABLED)
        MDIMAIN.lblkfc.Caption = IIf(IsNull(RSTCOMPANY!kfc_flag), "N", RSTCOMPANY!kfc_flag)
        If MDIMAIN.lblkfc.Caption = "Y" Then
            MDIMAIN.DTKFCSTART.Value = IIf(IsDate(RSTCOMPANY!KFCFROM_DATE), RSTCOMPANY!KFCFROM_DATE, Null)
            MDIMAIN.DTKFCEND.Value = IIf(IsDate(RSTCOMPANY!KFCTO_DATE), RSTCOMPANY!KFCTO_DATE, Null)
        Else
            MDIMAIN.DTKFCSTART.Value = Null
            MDIMAIN.DTKFCEND.Value = Null
        End If
        If RSTCOMPANY!STK_ADJ = "Y" Then MDIMAIN.MNUOPSTK.Visible = False
        MDIMAIN.lblub.Caption = IIf(IsNull(RSTCOMPANY!UB), "N", RSTCOMPANY!UB)
        MDIMAIN.barcode_profile.Caption = IIf(IsNull(RSTCOMPANY!barcode_profile), "", RSTCOMPANY!barcode_profile)
        MDIMAIN.LBLRT.Caption = IIf(IsNull(RSTCOMPANY!RTDISC), "", RSTCOMPANY!RTDISC)
        MDIMAIN.LBLWS.Caption = IIf(IsNull(RSTCOMPANY!WSDISC), "", RSTCOMPANY!WSDISC)
        MDIMAIN.lblvp.Caption = IIf(IsNull(RSTCOMPANY!VPDISC), "", RSTCOMPANY!VPDISC)
        MDIMAIN.LBLMRP.Caption = IIf(IsNull(RSTCOMPANY!MRPDISC), "", RSTCOMPANY!MRPDISC)
        MDIMAIN.LBLMRPPLUS.Caption = IIf(IsNull(RSTCOMPANY!mrp_plus), "", RSTCOMPANY!mrp_plus)
        MDIMAIN.LBLHSNSUM.Caption = IIf(IsNull(RSTCOMPANY!HSN_SUM), "", RSTCOMPANY!HSN_SUM)
        MDIMAIN.LblKFCNet.Caption = IIf(IsNull(RSTCOMPANY!KFCNET), "N", RSTCOMPANY!KFCNET)
        MDIMAIN.lblitemrepeat.Caption = IIf(IsNull(RSTCOMPANY!item_repeat), "N", RSTCOMPANY!item_repeat)
        PC_FLAG = IIf(IsNull(RSTCOMPANY!VS_FLAG), "N", RSTCOMPANY!VS_FLAG)
        SALESLT_FLAG = IIf(IsNull(RSTCOMPANY!LMT_FLAG), "N", RSTCOMPANY!LMT_FLAG)
        MINUS_BILL = IIf(IsNull(RSTCOMPANY!MINUS_BILL), "Y", RSTCOMPANY!MINUS_BILL)
        BARTEMPLATE = IIf(IsNull(RSTCOMPANY!BAR_TEMPLATE), "N", RSTCOMPANY!BAR_TEMPLATE)
        'BARFORMAT = IIf(IsNull(RSTCOMPANY!BAR_FORMAT), "N", RSTCOMPANY!BAR_FORMAT)
        MDIMAIN.lblmobwarn.Caption = IIf(IsNull(RSTCOMPANY!MOB_WARN_FLAG), "N", RSTCOMPANY!MOB_WARN_FLAG)
        D_PRINT = IIf(IsNull(RSTCOMPANY!D_PRINT) Or RSTCOMPANY!D_PRINT = "", "0", RSTCOMPANY!D_PRINT)
        PRNPETTYFLAG = IIf(RSTCOMPANY!PRN_PETTY_FLAG = "Y", True, False)
        bill_for = IIf(IsNull(RSTCOMPANY!BILL_FORMAT), "", RSTCOMPANY!BILL_FORMAT)
        MDIMAIN.lbldmpmini.Caption = IIf(IsNull(RSTCOMPANY!DMP_MINI), "N", RSTCOMPANY!DMP_MINI)
        
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
            MDIMAIN.CmdGST.Caption = "Sales"
            MDIMAIN.CmdRetailBill.Caption = "Sales"
            MDIMAIN.CmdRetailBill2.Caption = "Sales"
        End If
        
        If RSTCOMPANY!SHOP_RT = "Y" Then
            MDIMAIN.LBLSHOPRT.Caption = "Y"
        Else
            MDIMAIN.LBLSHOPRT.Caption = "N"
        End If
        If RSTCOMPANY!Zero_Warn = "N" Then
            ZERO_WARN_FLAG = False
        Else
            ZERO_WARN_FLAG = True
        End If
        If RSTCOMPANY!SAL_PROC = "Y" Then
            MDIMAIN.lblsalary.Caption = "Y"
            MDIMAIN.CmdStaff.Caption = "Salary Register"
        Else
            MDIMAIN.lblsalary.Caption = "N"
            MDIMAIN.CmdStaff.Caption = "Salary Expense (F12)"
        End If
        If RSTCOMPANY!CAT_PURCHASE = "Y" Then
            MDIMAIN.lblcategory.Caption = "Y"
        Else
            MDIMAIN.lblcategory.Caption = "N"
        End If
        
        MDIMAIN.LBLLABELNOS.Caption = IIf(IsNull(RSTCOMPANY!BAR_LABELS) Or RSTCOMPANY!BAR_LABELS = 0 Or RSTCOMPANY!BAR_LABELS = "", "1", RSTCOMPANY!BAR_LABELS)
        If RSTCOMPANY!RST_BILL = "Y" Then
            RstBill_Flag = "Y"
        Else
            RstBill_Flag = "N"
        End If
        If RSTCOMPANY!PRICE_SPLIT = "Y" Then
            MDIMAIN.lblPriceSplit.Caption = "Y"
        Else
            MDIMAIN.lblPriceSplit.Caption = "N"
        End If
        If RSTCOMPANY!PER_PURCHASE = "Y" Then
            MDIMAIN.lblPerPurchase.Caption = "Y"
        Else
            MDIMAIN.lblPerPurchase.Caption = "N"
        End If
        If RSTCOMPANY!SCHEME_OPT = "1" Then
            scheme_option = "1"
        Else
            scheme_option = "0"
        End If
        MDIMAIN.Caption = "EzBiz Inventory - Financial Year " & Trim(txtyear.Text) & " - " & Trim(lblfinyear.Caption) & " (" & Trim(MDIMAIN.StatusBar.Panels(5).Text) & ")"
        On Error Resume Next
        If err.Number = 545 Then
            MDIMAIN.Picture = LoadPicture("")
        Else
             MDIMAIN.Picture = LoadPicture(RSTCOMPANY!COMP_LOGO)
        End If
        On Error GoTo ERRHAND
    Else
       
        Dim RSTCOMPANY2 As ADODB.Recordset
        Set RSTCOMPANY2 = New ADODB.Recordset
        RSTCOMPANY2.Open "Select * From COMPINFO WHERE COMP_CODE = '001' ORDER BY FIN_YR DESC", db, adOpenStatic, adLockReadOnly
        If Not (RSTCOMPANY2.EOF And RSTCOMPANY2.BOF) Then
            RSTCOMPANY.AddNew
            RSTCOMPANY!COMP_CODE = "001"
            RSTCOMPANY!CUST_CODE = RSTCOMPANY2!CUST_CODE
            RSTCOMPANY!SERVER_ADD = RSTCOMPANY2!SERVER_ADD
            RSTCOMPANY!CLOUD_TYPE = RSTCOMPANY2!CLOUD_TYPE
            RSTCOMPANY!FIN_YR = txtyear.Text
            RSTCOMPANY!MULTI_FLAG = RSTCOMPANY2!MULTI_FLAG
            RSTCOMPANY!REMOTE_FLAG = RSTCOMPANY2!REMOTE_FLAG
            RSTCOMPANY!COMP_NAME = RSTCOMPANY2!COMP_NAME
            RSTCOMPANY!Address = RSTCOMPANY2!Address
            RSTCOMPANY!TEL_NO = RSTCOMPANY2!TEL_NO
            RSTCOMPANY!FAX_NO = RSTCOMPANY2!FAX_NO
            RSTCOMPANY!EMAIL_ADD = RSTCOMPANY2!EMAIL_ADD
            RSTCOMPANY!KGST = RSTCOMPANY2!KGST
            RSTCOMPANY!CST = RSTCOMPANY2!CST
            RSTCOMPANY!SCODE = RSTCOMPANY2!SCODE
            RSTCOMPANY!SNAME = RSTCOMPANY2!SNAME
            RSTCOMPANY!SAL_PROC = RSTCOMPANY2!SAL_PROC
            RSTCOMPANY!CAT_PURCHASE = RSTCOMPANY2!CAT_PURCHASE
            RSTCOMPANY!BAR_LABELS = RSTCOMPANY2!BAR_LABELS
            RSTCOMPANY!RST_BILL = RSTCOMPANY2!RST_BILL
            RSTCOMPANY!PRICE_SPLIT = RSTCOMPANY2!PRICE_SPLIT
            RSTCOMPANY!PER_PURCHASE = RSTCOMPANY2!PER_PURCHASE
            RSTCOMPANY!DL_NO = RSTCOMPANY2!DL_NO
            RSTCOMPANY!ML_NO = RSTCOMPANY2!ML_NO
            RSTCOMPANY!HO_NAME = RSTCOMPANY2!HO_NAME
            RSTCOMPANY!PINCODE = RSTCOMPANY2!PINCODE
            RSTCOMPANY!auth_key = RSTCOMPANY2!auth_key
            RSTCOMPANY!INV_MSGS = RSTCOMPANY2!INV_MSGS
            RSTCOMPANY!PREFIX_8B = RSTCOMPANY2!PREFIX_8B
            RSTCOMPANY!SUFIX_8B = RSTCOMPANY2!SUFIX_8B
            RSTCOMPANY!PREFIX_8 = RSTCOMPANY2!PREFIX_8
            RSTCOMPANY!SUFIX_8 = RSTCOMPANY2!SUFIX_8
            RSTCOMPANY!PREFIX_8V = RSTCOMPANY2!PREFIX_8V
            RSTCOMPANY!SUFIX_8V = RSTCOMPANY2!SUFIX_8V
            RSTCOMPANY!PREFIX_TR = RSTCOMPANY2!PREFIX_TR
            RSTCOMPANY!SUFIX_TR = RSTCOMPANY2!SUFIX_TR
            RSTCOMPANY!VEHICLE = RSTCOMPANY2!VEHICLE
            RSTCOMPANY!CGST = RSTCOMPANY2!CGST
            RSTCOMPANY!SGST = RSTCOMPANY2!SGST
            RSTCOMPANY!IGST = RSTCOMPANY2!IGST
            RSTCOMPANY!RTDISC = RSTCOMPANY2!RTDISC
            RSTCOMPANY!WSDISC = RSTCOMPANY2!WSDISC
            RSTCOMPANY!VPDISC = RSTCOMPANY2!VPDISC
            RSTCOMPANY!MRPDISC = RSTCOMPANY2!MRPDISC
            RSTCOMPANY!HSN_SUM = RSTCOMPANY2!HSN_SUM
            RSTCOMPANY!GST_FLAG = RSTCOMPANY2!GST_FLAG
            RSTCOMPANY!Copy_8B = RSTCOMPANY2!Copy_8B
            RSTCOMPANY!Copy_8 = RSTCOMPANY2!Copy_8
            RSTCOMPANY!Copy_8V = RSTCOMPANY2!Copy_8V
            RSTCOMPANY!Copy_TR = RSTCOMPANY2!Copy_TR
            RSTCOMPANY!NSPT = RSTCOMPANY2!NSPT
            RSTCOMPANY!ALL_PRN = RSTCOMPANY2!ALL_PRN
            RSTCOMPANY!UB = RSTCOMPANY2!UB
            RSTCOMPANY!ONLINE_BILL = RSTCOMPANY2!ONLINE_BILL
            RSTCOMPANY!BILL_FORMAT = RSTCOMPANY2!BILL_FORMAT
            RSTCOMPANY!STK_ADJ = RSTCOMPANY2!STK_ADJ
            RSTCOMPANY!EC = RSTCOMPANY2!EC
            RSTCOMPANY!DMP_FLAG = RSTCOMPANY2!DMP_FLAG
            RSTCOMPANY!LINE_SPACE = RSTCOMPANY2!LINE_SPACE
            RSTCOMPANY!DMPTH_FLAG = RSTCOMPANY2!DMPTH_FLAG
            RSTCOMPANY!TAX_FLAG = RSTCOMPANY2!TAX_FLAG
            RSTCOMPANY!KFCNET = RSTCOMPANY2!KFCNET
            RSTCOMPANY!TAXWRN_FLAG = RSTCOMPANY2!TAXWRN_FLAG
            RSTCOMPANY!GSTWRN_FLAG = RSTCOMPANY2!GSTWRN_FLAG
            RSTCOMPANY!ITEM_WARN = RSTCOMPANY2!ITEM_WARN
            RSTCOMPANY!AMC_FLAG = RSTCOMPANY2!AMC_FLAG
            RSTCOMPANY!HOLD_THERMAL_FLAG = RSTCOMPANY2!HOLD_THERMAL_FLAG
            RSTCOMPANY!FORM_62 = RSTCOMPANY2!FORM_62
            RSTCOMPANY!DUP_FLAG = RSTCOMPANY2!DUP_FLAG
            RSTCOMPANY!ROUND_FLAG = RSTCOMPANY2!ROUND_FLAG
            RSTCOMPANY!PREVIEW_FLAG = RSTCOMPANY2!PREVIEW_FLAG
            RSTCOMPANY!PREVIEWTH_FLAG = RSTCOMPANY2!PREVIEWTH_FLAG
            RSTCOMPANY!CODE_FLAG = RSTCOMPANY2!CODE_FLAG
            RSTCOMPANY!DISC_FLAG = RSTCOMPANY2!DISC_FLAG
            RSTCOMPANY!BARCODE_FLAG = RSTCOMPANY2!BARCODE_FLAG
            RSTCOMPANY!barcode_profile = RSTCOMPANY2!barcode_profile
            RSTCOMPANY!OSSR_FLAG = RSTCOMPANY2!OSSR_FLAG
            RSTCOMPANY!OSB2C_FLAG = RSTCOMPANY2!OSB2C_FLAG
            RSTCOMPANY!OSB2B_FLAG = RSTCOMPANY2!OSB2B_FLAG
            RSTCOMPANY!OSPTY_FLAG = RSTCOMPANY2!OSPTY_FLAG
            RSTCOMPANY!REMOVE_UBILL = RSTCOMPANY2!REMOVE_UBILL
            RSTCOMPANY!MRP_DISC = RSTCOMPANY2!MRP_DISC
            RSTCOMPANY!EXP_ENABLED = RSTCOMPANY2!EXP_ENABLED
            RSTCOMPANY!TERMS_FLAG = RSTCOMPANY!TERMS_FLAG
            RSTCOMPANY!Terms1 = RSTCOMPANY2!Terms1
            RSTCOMPANY!Terms2 = RSTCOMPANY2!Terms2
            RSTCOMPANY!Terms3 = RSTCOMPANY2!Terms3
            RSTCOMPANY!Terms4 = RSTCOMPANY2!Terms4
            RSTCOMPANY!INV_TERMS = RSTCOMPANY2!INV_TERMS
            RSTCOMPANY!bank_details = RSTCOMPANY2!bank_details
            RSTCOMPANY!PAN_NO = RSTCOMPANY2!PAN_NO
            RSTCOMPANY!SHOP_RT = RSTCOMPANY2!SHOP_RT
            RSTCOMPANY!hide_spec = RSTCOMPANY2!hide_spec
            RSTCOMPANY!hide_pr_name = RSTCOMPANY2!hide_pr_name
            RSTCOMPANY!hide_wrnty = RSTCOMPANY2!hide_wrnty
            RSTCOMPANY!hide_serial = RSTCOMPANY2!hide_serial
            RSTCOMPANY!hide_mrp = RSTCOMPANY2!hide_mrp
            RSTCOMPANY!Zero_Warn = RSTCOMPANY2!Zero_Warn
            RSTCOMPANY!billtype_flag = RSTCOMPANY2!billtype_flag
            RSTCOMPANY!hide_expiry = RSTCOMPANY2!hide_expiry
            RSTCOMPANY!hide_free = RSTCOMPANY2!hide_free
            RSTCOMPANY!hide_disc = RSTCOMPANY2!hide_disc
            RSTCOMPANY!hide_terms = RSTCOMPANY2!hide_terms
            RSTCOMPANY!hide_deliver = RSTCOMPANY2!hide_deliver
            RSTCOMPANY!mrp_plus = RSTCOMPANY2!mrp_plus
            RSTCOMPANY!billtype_flag = RSTCOMPANY2!billtype_flag
            RSTCOMPANY!kfc_flag = RSTCOMPANY2!kfc_flag
            If IsDate(RSTCOMPANY2!KFCFROM_DATE) Then
                RSTCOMPANY!KFCFROM_DATE = Format(RSTCOMPANY2!KFCFROM_DATE, "DD/MM/YYYY")
            End If
            If IsDate(RSTCOMPANY2!KFCTO_DATE) Then
                RSTCOMPANY!KFCTO_DATE = Format(RSTCOMPANY2!KFCTO_DATE, "DD/MM/YYYY")
            End If
            RSTCOMPANY!item_repeat = RSTCOMPANY2!item_repeat
            RSTCOMPANY!VS_FLAG = RSTCOMPANY2!VS_FLAG
            RSTCOMPANY!LMT_FLAG = RSTCOMPANY2!LMT_FLAG
            RSTCOMPANY!MINUS_BILL = RSTCOMPANY2!MINUS_BILL
            RSTCOMPANY!BAR_TEMPLATE = RSTCOMPANY2!BAR_TEMPLATE
            'RSTCOMPANY!BAR_FORMAT = RSTCOMPANY2!BAR_FORMAT
            RSTCOMPANY!MOB_WARN_FLAG = RSTCOMPANY2!MOB_WARN_FLAG
            RSTCOMPANY!CLCODE = RSTCOMPANY2!CLCODE
            RSTCOMPANY!PCODE = RSTCOMPANY2!PCODE
            RSTCOMPANY!SCHEME_OPT = RSTCOMPANY2!SCHEME_OPT
            RSTCOMPANY.Update
            
            customercode = IIf(IsNull(RSTCOMPANY!CUST_CODE), "", RSTCOMPANY!CUST_CODE)
            serveraddress = IIf(IsNull(RSTCOMPANY!SERVER_ADD), "", RSTCOMPANY!SERVER_ADD)
            REMOTEAPP = IIf(IsNull(RSTCOMPANY!CLOUD_TYPE), 0, RSTCOMPANY!CLOUD_TYPE)
            If RSTCOMPANY!REMOTE_FLAG = "Y" Then
                remoteflag = True
                MDIMAIN.mnusync.Visible = True
            Else
                remoteflag = False
                MDIMAIN.mnusync.Visible = False
            End If
            DUPCODE = IIf(IsNull(RSTCOMPANY!PCODE), "", RSTCOMPANY!PCODE)
            MDIMAIN.lblec.Caption = IIf(IsNull(RSTCOMPANY!EC), "", RSTCOMPANY!EC)
            MDIMAIN.StatusBar.Panels(5).Text = IIf(IsNull(RSTCOMPANY!COMP_NAME), "", RSTCOMPANY!COMP_NAME)
            MDIMAIN.StatusBar.Panels(8).Text = IIf(IsNull(RSTCOMPANY!DMP_FLAG), "N", RSTCOMPANY!DMP_FLAG)
            MDIMAIN.LBLSPACE.Caption = IIf(IsNull(RSTCOMPANY!LINE_SPACE), "N", RSTCOMPANY!LINE_SPACE)
            MDIMAIN.LBLDMPTHERMAL.Caption = IIf(IsNull(RSTCOMPANY!DMPTH_FLAG), "N", RSTCOMPANY!DMPTH_FLAG)
            MDIMAIN.StatusBar.Panels(9).Text = IIf(IsNull(RSTCOMPANY!DUP_FLAG), "N", RSTCOMPANY!DUP_FLAG)
            Roundflag = IIf(IsNull(RSTCOMPANY!ROUND_FLAG) Or RSTCOMPANY!ROUND_FLAG <> "N", True, False)
            BATCH_DISPLAY = IIf(IsNull(RSTCOMPANY!BATCH_FLAG) Or RSTCOMPANY!BATCH_FLAG <> "Y", False, True)
            MDIMAIN.StatusBar.Panels(10).Text = IIf(IsNull(RSTCOMPANY!Copy_8B), "1", RSTCOMPANY!Copy_8B)
            MDIMAIN.StatusBar.Panels(11).Text = IIf(IsNull(RSTCOMPANY!Copy_8), "1", RSTCOMPANY!Copy_8)
            MDIMAIN.StatusBar.Panels(12).Text = IIf(IsNull(RSTCOMPANY!Copy_8V), "1", RSTCOMPANY!Copy_8V)
            MDIMAIN.LBLTRCopy.Caption = IIf(IsNull(RSTCOMPANY!Copy_TR), "1", RSTCOMPANY!Copy_TR)
            MDIMAIN.StatusBar.Panels(13).Text = IIf(IsNull(RSTCOMPANY!PREVIEW_FLAG), "N", RSTCOMPANY!PREVIEW_FLAG)
            MDIMAIN.LBLTHPREVIEW.Caption = IIf(IsNull(RSTCOMPANY!PREVIEWTH_FLAG), "N", RSTCOMPANY!PREVIEWTH_FLAG)
            MDIMAIN.StatusBar.Panels(14).Text = IIf(IsNull(RSTCOMPANY!TAX_FLAG), "N", RSTCOMPANY!TAX_FLAG)
            MDIMAIN.StatusBar.Panels(15).Text = IIf(IsNull(RSTCOMPANY!CODE_FLAG), "N", RSTCOMPANY!CODE_FLAG)
            MDIMAIN.StatusBar.Panels(16).Text = IIf(IsNull(RSTCOMPANY!DISC_FLAG), "N", RSTCOMPANY!DISC_FLAG)
            MDIMAIN.StatusBar.Panels(6).Text = IIf(IsNull(RSTCOMPANY!BARCODE_FLAG), "N", RSTCOMPANY!BARCODE_FLAG)
            MDIMAIN.LBLTAXWARN.Caption = IIf(IsNull(RSTCOMPANY!TAXWRN_FLAG), "N", RSTCOMPANY!TAXWRN_FLAG)
            MDIMAIN.LBLAMC.Caption = IIf(IsNull(RSTCOMPANY!AMC_FLAG), "N", RSTCOMPANY!AMC_FLAG)
            hold_thermal = IIf(RSTCOMPANY!HOLD_THERMAL_FLAG = "Y", True, False)
            MDIMAIN.LBLGSTWRN.Caption = IIf(IsNull(RSTCOMPANY!GSTWRN_FLAG), "Y", RSTCOMPANY!GSTWRN_FLAG)
            MDIMAIN.LBLITMWRN.Caption = IIf(IsNull(RSTCOMPANY!ITEM_WARN), "Y", RSTCOMPANY!ITEM_WARN)
            MDIMAIN.lblform62.Caption = IIf(IsNull(RSTCOMPANY!FORM_62), "N", RSTCOMPANY!FORM_62)
            MDIMAIN.lblgst.Caption = IIf(IsNull(RSTCOMPANY!GST_FLAG), "R", RSTCOMPANY!GST_FLAG)
            MDIMAIN.lblnostock.Caption = IIf(IsNull(RSTCOMPANY!NSPT), "N", RSTCOMPANY!NSPT)
            MDIMAIN.lblprnall.Caption = IIf(IsNull(RSTCOMPANY!ALL_PRN), "N", RSTCOMPANY!ALL_PRN)
            MDIMAIN.lblRemoveUbill.Caption = IIf(IsNull(RSTCOMPANY!REMOVE_UBILL), "Y", RSTCOMPANY!REMOVE_UBILL)
            MRPDISC_FLAG = IIf(IsNull(RSTCOMPANY!MRP_DISC), "N", RSTCOMPANY!MRP_DISC)
            MDIMAIN.LBLSTATE.Caption = IIf(IsNull(RSTCOMPANY!SCODE) Or RSTCOMPANY!SCODE = 0 Or RSTCOMPANY!SCODE = "", "32", RSTCOMPANY!SCODE)
            MDIMAIN.LBLSTATENAME.Caption = IIf(IsNull(RSTCOMPANY!SNAME) Or RSTCOMPANY!SNAME = "", "KL", RSTCOMPANY!SNAME)
            If Trim(MDIMAIN.LBLSTATE.Caption) = "" Then MDIMAIN.LBLSTATE.Caption = "32"
            If Val(MDIMAIN.LBLSTATE.Caption) = 0 Then MDIMAIN.LBLSTATE.Caption = "32"
            If Len(MDIMAIN.LBLSTATE.Caption) <> 2 Then MDIMAIN.LBLSTATE.Caption = "32"
            If Val(MDIMAIN.LBLSTATE.Caption) = 32 Then MDIMAIN.LBLSTATENAME.Caption = "KL"
            If RSTCOMPANY!STK_ADJ = "Y" Then MDIMAIN.MNUOPSTK.Visible = False
            
            MDIMAIN.lblExpEnable.Caption = IIf(IsNull(RSTCOMPANY!EXP_ENABLED), "N", RSTCOMPANY!EXP_ENABLED)
            MDIMAIN.lblkfc.Caption = IIf(IsNull(RSTCOMPANY!kfc_flag), "N", RSTCOMPANY!kfc_flag)
            If MDIMAIN.lblkfc.Caption = "Y" Then
                MDIMAIN.DTKFCSTART.Value = IIf(IsDate(RSTCOMPANY!KFCFROM_DATE), RSTCOMPANY!KFCFROM_DATE, Null)
                MDIMAIN.DTKFCEND.Value = IIf(IsDate(RSTCOMPANY!KFCTO_DATE), RSTCOMPANY!KFCTO_DATE, Null)
            Else
                MDIMAIN.DTKFCSTART.Value = Null
                MDIMAIN.DTKFCEND.Value = Null
            End If
            MDIMAIN.lblub.Caption = IIf(IsNull(RSTCOMPANY!UB), "N", RSTCOMPANY!UB)
            MDIMAIN.barcode_profile.Caption = IIf(IsNull(RSTCOMPANY!barcode_profile), "", RSTCOMPANY!barcode_profile)
            MDIMAIN.LBLRT.Caption = IIf(IsNull(RSTCOMPANY!RTDISC), "", RSTCOMPANY!RTDISC)
            MDIMAIN.LBLWS.Caption = IIf(IsNull(RSTCOMPANY!WSDISC), "", RSTCOMPANY!WSDISC)
            MDIMAIN.lblvp.Caption = IIf(IsNull(RSTCOMPANY!VPDISC), "", RSTCOMPANY!VPDISC)
            MDIMAIN.LBLMRP.Caption = IIf(IsNull(RSTCOMPANY!MRPDISC), "", RSTCOMPANY!MRPDISC)
            MDIMAIN.LBLMRPPLUS.Caption = IIf(IsNull(RSTCOMPANY!mrp_plus), "", RSTCOMPANY!mrp_plus)
            MDIMAIN.LBLHSNSUM.Caption = IIf(IsNull(RSTCOMPANY!HSN_SUM), "", RSTCOMPANY!HSN_SUM)
            MDIMAIN.LblKFCNet.Caption = IIf(IsNull(RSTCOMPANY!KFCNET), "N", RSTCOMPANY!KFCNET)
            MDIMAIN.lblitemrepeat.Caption = IIf(IsNull(RSTCOMPANY!item_repeat), "N", RSTCOMPANY!item_repeat)
            PC_FLAG = IIf(IsNull(RSTCOMPANY!VS_FLAG), "N", RSTCOMPANY!VS_FLAG)
            SALESLT_FLAG = IIf(IsNull(RSTCOMPANY!LMT_FLAG), "N", RSTCOMPANY!LMT_FLAG)
            MINUS_BILL = IIf(IsNull(RSTCOMPANY!MINUS_BILL), "Y", RSTCOMPANY!MINUS_BILL)
            BARTEMPLATE = IIf(IsNull(RSTCOMPANY!BAR_TEMPLATE), "N", RSTCOMPANY!BAR_TEMPLATE)
            'BARFORMAT = IIf(IsNull(RSTCOMPANY!BAR_FORMAT), "N", RSTCOMPANY!BAR_FORMAT)
            MDIMAIN.lblmobwarn.Caption = IIf(IsNull(RSTCOMPANY!MOB_WARN_FLAG), "N", RSTCOMPANY!MOB_WARN_FLAG)
            D_PRINT = IIf(IsNull(RSTCOMPANY!D_PRINT) Or RSTCOMPANY!D_PRINT = "", "0", RSTCOMPANY!D_PRINT)
            PRNPETTYFLAG = IIf(RSTCOMPANY!PRN_PETTY_FLAG = "Y", True, False)
            bill_for = IIf(IsNull(RSTCOMPANY!BILL_FORMAT), "", RSTCOMPANY!BILL_FORMAT)
            MDIMAIN.lbldmpmini.Caption = IIf(IsNull(RSTCOMPANY!DMP_MINI), "N", RSTCOMPANY!DMP_MINI)
            If RSTCOMPANY!billtype_flag = "Y" Then
                bill_type_flag = True
            Else
                bill_type_flag = False
            End If
            If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
                MDIMAIN.CmdGST.Caption = "Sales"
                MDIMAIN.CmdRetailBill.Caption = "Sales"
                MDIMAIN.CmdRetailBill2.Caption = "Sales"
            End If
            
            If RSTCOMPANY!SHOP_RT = "Y" Then
                MDIMAIN.LBLSHOPRT.Caption = "Y"
            Else
                MDIMAIN.LBLSHOPRT.Caption = "N"
            End If
            If RSTCOMPANY!Zero_Warn = "N" Then
                ZERO_WARN_FLAG = False
            Else
                ZERO_WARN_FLAG = True
            End If
            If RSTCOMPANY!SCHEME_OPT = "1" Then
                scheme_option = "1"
            Else
                scheme_option = "0"
            End If
            MDIMAIN.Caption = "EzBiz Inventory - Financial Year " & Trim(txtyear.Text) & " - " & Trim(lblfinyear.Caption) & " (" & Trim(MDIMAIN.StatusBar.Panels(5).Text) & ")"
            On Error Resume Next
            If err.Number = 545 Then
                MDIMAIN.Picture = LoadPicture("")
            Else
                 MDIMAIN.Picture = LoadPicture(RSTCOMPANY!COMP_LOGO)
            End If
            On Error GoTo ERRHAND
        End If
        RSTCOMPANY2.Close
        Set RSTCOMPANY2 = Nothing
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
'    If (Year(Date) >= 2021 And Month(Date) >= 5 And Day(Date) >= 18) Then
'        'db.Execute "Delete from Users"
'        db.Execute "ALTER TABLE `users` CHANGE `USER_NAME` `USER_NME` VARCHAR(30) CHARACTER SET utf8 COLLATE utf8_general_ci NOT NULL"
'        MsgBox "Abnormal Program Termination....", vbCritical, "EzBiz"
'        db.Close
'        Set db = Nothing
'
'        dbprint.Close
'        Set dbprint = Nothing
'    End If

    
    
'    StrFileName = "C:\User\EX"
'    Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
'    Set ObjFile = ObjFso.OpenTextFile(StrFileName)  'Reading from the file
'    MDIMAIN.StatusBar.Panels(7).Text = ObjFile.ReadLine & "\"
        
    
    If FileExists(App.Path & "\print.txt") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\print.txt")  'Reading from the file
        MDIMAIN.lblprint.Caption = ObjFile.ReadLine
        Set ObjFso = Nothing
        Set ObjFile = Nothing
    End If
    
    DMPrint = ""
    If FileExists(App.Path & "\dmp.txt") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\dmp.txt")  'Reading from the file
        DMPrint = ObjFile.ReadLine
        Set ObjFso = Nothing
        Set ObjFile = Nothing
    End If
    
    DMPrintA4 = ""
    If FileExists(App.Path & "\dmpA4.txt") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\dmpA4.txt")  'Reading from the file
        DMPrintA4 = ObjFile.ReadLine
        Set ObjFso = Nothing
        Set ObjFile = Nothing
    End If
    
    If DMPrintA4 = "" Then DMPrintA4 = DMPrint
    
    BarPrint = ""
    If FileExists(App.Path & "\barprint.txt") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\barprint.txt")  'Reading from the file
        BarPrint = ObjFile.ReadLine
        Set ObjFso = Nothing
        Set ObjFile = Nothing
    End If
    'On Error GoTo Errhand
    
    
    rs!LAST_LOGIN = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss")
    rs!LAST_LOGOUT = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss")
    rs.Update
    
End If
'MsgBox "User Name: " & rs!USER_NAME & Chr(13) & "Last Login : " & IIf(IsNull(rs!LAST_LOGIN), "", rs!LAST_LOGIN)

Dim RSTITEMMAST As ADODB.Recordset
Set RSTITEMMAST = New ADODB.Recordset
RSTITEMMAST.Open "SELECT * FROM cont_mast WHERE CONT_NAME = '" & system_name & "'", db, adOpenStatic, adLockOptimistic, adCmdText
If (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
    RSTITEMMAST.AddNew
    RSTITEMMAST!CONT_NAME = system_name
    RSTITEMMAST!CREATE_DATE = Format(Date, "DD/MM/YYYY")
    RSTITEMMAST!C_USER_ID = rs!USER_ID
    RSTITEMMAST!M_USER_ID = rs!USER_ID
    RSTITEMMAST!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
    RSTITEMMAST!OP_AMT = 0
    RSTITEMMAST.Update
End If
RSTITEMMAST.Close
Set RSTITEMMAST = Nothing

MDIMAIN.Show
MDIMAIN.DTFROM.Value = "01/04/" & txtyear.Text

If remoteflag = True Then
    If MsgBox("Do you want to autmatically sync with cloud", vbYesNo + vbDefaultButton2, "EzBiz") = vbYes Then
        If REMOTEAPP = 1 Then
            Call export_db
        Else
            Call export_db2
        End If
    End If
End If
Unload Me
Exit Sub

ERRHAND:
    If err.Number = 3709 Then
        ' "USER NOT FOUND", vbOKOnly, "EzBiz"
        MsgBox err.Description, , "EzBiz"
        End
    ElseIf err.Number = 3265 Then
        'MsgBox "Abnormal Program Termination.", vbCritical, "EzBiz"
        MsgBox err.Description, , "EzBiz"
        db.Close
        Set db = Nothing
        
'        dbprint.Close
'        Set dbprint = Nothing
        End
    ElseIf err.Number = -2147217900 Then
        'MsgBox "Crxdtl.dll missing. Abnormal Program Termination..", vbCritical, "EzBiz"
        MsgBox err.Description, , "EzBiz"
        db.Close
        Set db = Nothing
        
'        dbprint.Close
'        Set dbprint = Nothing
        End
    End If
    rs.Close
    db.Close
    Set db = Nothing
    
'    dbprint.Close
'    Set dbprint = Nothing
    'MsgBox GetUniqueCode
    'MsgBox Err.Description, vbOKOnly, "EzBiz"
    If err.Number = -2147467259 Then
        MsgBox err.Description, , "EzBiz"
        'MsgBox "DATABASE CORRUPTED !!! PLEASE CONTACT PROGRAM VENDOR", vbOKOnly, "LOGIN"
        End
    Else
        MsgBox err.Description, , "EzBiz"
        'MsgBox "Crxdtl.dll missing. Abnormal Program Termination...", vbCritical, "EzBiz"
    End If
    End
    
End Sub

Private Sub cmdRestore_Click()
    
End Sub

Private Sub Form_Activate()
If Month(Date) < 4 Then
    txtyear.Text = Format(Year(Date) - 1, "####")
Else
    txtyear.Text = Format(Year(Date), "####")
End If
'lblfinyear.Caption = Val(txtyear.Text) + 1
'Place form in center of the screen
cetre Me
End Sub

Private Sub Form_Load()
    Dim sql As String
    
    'sql = "select * from USERS where LEVEL='0' "
    sql = "select * from USERS"
    rs.Open sql, db, adOpenKeyset, adLockPessimistic
    If rs.RecordCount = 1 Then
        txtPassword.TabIndex = 0
        txtLogin.Text = IIf(IsNull(rs!USER_NAME), "", rs!USER_NAME)
    Else
        txtLogin.TabIndex = 0
        'txtpassword.SetFocus
    End If
    rs.Close
    
    sql = "select * from compinfo WHERE COMP_CODE = '001' ORDER BY FIN_YR DESC"
    rs.Open sql, db, adOpenKeyset, adLockPessimistic
    If rs.RecordCount >= 1 Then
        If rs!MULTI_FLAG = "Y" Then
            LBLCOMP.Visible = True
            CmbDB.Visible = True
            Call GetDBNames
        Else
            LBLCOMP.Visible = False
            CmbDB.Visible = False
        End If
    End If
    rs.Close
    
    
    
    
''    'Dim i As Long
''    Dim RSTITEMMAST As ADODB.Recordset
''    'Dim strConn As String
''    Dim strConnect As String
''    Dim DBPwd As String
''
'''    On Error GoTo eRRHAND
'''    Set db = New ADODB.Connection
'''    DBPath = MDIMAIN.StatusBar.Panels(6).Text
'''    DBPwd = "###DATABASE%%%RET"
'''    'DBPwd = "DATA_RET"
'''    strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & DBPath & ";Jet OLEDB:Database Password=" & DBPwd
'''    db.Open strConnect
'''
'''    db.CursorLocation = adUseClient
''
''    'i = 9
    
    
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub mnufix_Click()
    If MsgBox("Are you sure?", vbYesNo, "Fix Error") = vbNo Then Exit Sub
    db.Execute "CHECK TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
    db.Execute "CHECK TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
    db.Execute "CHECK TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
    db.Execute "CHECK TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
    db.Execute "CHECK TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
    db.Execute "CHECK TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
    db.Execute "CHECK TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
    db.Execute "CHECK TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
    db.Execute "CHECK TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
    db.Execute "CHECK TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
    
    db.Execute "OPTIMIZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
    db.Execute "OPTIMIZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
    db.Execute "OPTIMIZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
    db.Execute "OPTIMIZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
    db.Execute "OPTIMIZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
    db.Execute "OPTIMIZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
    db.Execute "OPTIMIZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
    db.Execute "OPTIMIZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
    db.Execute "OPTIMIZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
    db.Execute "OPTIMIZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
    
    db.Execute "REPAIR TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
    db.Execute "REPAIR TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
    db.Execute "REPAIR TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
    db.Execute "REPAIR TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
    db.Execute "REPAIR TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
    db.Execute "REPAIR TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
    db.Execute "REPAIR TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
    db.Execute "REPAIR TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
    db.Execute "REPAIR TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
    db.Execute "REPAIR TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
    
    db.Execute "ANALYZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
    db.Execute "ANALYZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
    db.Execute "ANALYZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
    db.Execute "ANALYZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
    db.Execute "ANALYZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
    db.Execute "ANALYZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
    db.Execute "ANALYZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
    db.Execute "ANALYZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
    db.Execute "ANALYZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
    db.Execute "ANALYZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
    
    db.Execute "FLUSH TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
    db.Execute "FLUSH TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
    db.Execute "FLUSH TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
    db.Execute "FLUSH TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
    db.Execute "FLUSH TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
    db.Execute "FLUSH TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
    db.Execute "FLUSH TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
    db.Execute "FLUSH TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
    db.Execute "FLUSH TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
    db.Execute "FLUSH TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
    
    On Error Resume Next
    db.Execute "CHECK TABLE `astmast`, `astrxfile`, `astrxmast`, "
    db.Execute "OPTIMIZE TABLE `astmast`, `astrxfile`, `astrxmast`, "
    db.Execute "REPAIR TABLE `astmast`, `astrxfile`, `astrxmast`, "
    
    db.Execute "CHECK TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
    db.Execute "OPTIMIZE TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
    db.Execute "REPAIR TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
    
    On Error GoTo ERRHAND
        
    db.Execute "Update itemmast set close_qty = 0 where category = 'SELF' OR category = 'SERVICE CHARGE' "
    db.Execute "Update itemmast set check_flag = 'V' "
    db.Execute "Update rtrxfile set check_flag = 'V' "
    db.Execute "Update itemmast set close_qty = 0 where isnull(close_qty) "
    db.Execute "Update rtrxfile set bal_qty = 0 where isnull(bal_qty) "
    db.Execute "Update rtrxfile set category = '' where isnull(category) "
    db.Execute "Update rtrxfile set ref_no = '' where isnull(ref_no) "
    db.Execute "Update rtrxfile set TRX_GODOWN = '' where isnull(TRX_GODOWN) "
    db.Execute "Update itemmast set category = '' where isnull(category) "
    db.Execute "Update itemmast set REMARKS = '' where isnull(REMARKS) "
    db.Execute "Update itemmast set BIN_LOCATION = '' where isnull(BIN_LOCATION) "
    db.Execute "Update itemmast set ITEM_SPEC = '' where isnull(ITEM_SPEC) "
    MsgBox "Success", vbOKOnly, "EzBiz"
    Exit Sub
ERRHAND:
    MsgBox err.Description
End Sub

Private Sub mnurestore_Click()
    FrmRestore.Show
    FrmRestore.SetFocus
End Sub

Private Sub txtLogin_GotFocus()
    txtLogin.SelStart = 0
    txtLogin.SelLength = Len(txtLogin.Text)
    txtLogin.BackColor = &H98F3C1
End Sub
Private Sub txtLogin_KeyDown(KeyCode As Integer, Shift As Integer)
     Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If txtLogin.Text = "" Then Exit Sub
            txtPassword.SetFocus
    End Select
End Sub

Private Sub txtLogin_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack, vbKeyA To vbKeyZ, Asc("a") To Asc("z"), Asc("."), Asc("-"), Asc(" ")
            KeyAscii = Asc(Chr(KeyAscii))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtLogin_LostFocus()
    txtLogin.BackColor = vbWhite
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword.Text)
    txtPassword.BackColor = &H98F3C1
End Sub

Private Sub txtpassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn
            If txtPassword.Text = "" Then Exit Sub
            cmdLogin_Click
        Case vbKeyTab
            txtyear.SetFocus
        Case vbKeyEscape
            txtLogin.SetFocus
    End Select
End Sub

Private Sub txtPassword_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
    End Select
End Sub

Private Function deletedatabase()

    Dim RSTTRXFILE As ADODB.Recordset
    
    db.Execute "delete FROM ATRXFILE"
    db.Execute "delete FROM ATRXSUB"
    db.Execute "delete FROM BANKCODE"
    db.Execute "delete FROM BANKLETTERS"
    db.Execute "delete FROM BONUSMAST"
    db.Execute "delete FROM CANCINV"
    db.Execute "delete FROM CHQMAST"
    db.Execute "delete FROM DAMAGED"
    db.Execute "delete FROM FQTYLIST"
    db.Execute "delete FROM ORDISSUE"
    db.Execute "delete FROM ORDSUB"
    db.Execute "delete FROM USERS"
    db.Execute "delete FROM POMAST"
    db.Execute "delete FROM POSUB"
    db.Execute "delete FROM PRICETABLE"
    db.Execute "delete FROM QTNMAST"
    db.Execute "delete FROM QTNSUB"
    db.Execute "delete FROM REORDER"
    'db.Execute "delete FROM RTRXFILE"
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "Select * From ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
    Do Until RSTTRXFILE.EOF
        RSTTRXFILE!BIN_LOCATION = Mid(RSTTRXFILE!ITEM_NAME, 1, 1)
        RSTTRXFILE!SALES_TAX = 0
        RSTTRXFILE!CST = 0
        RSTTRXFILE!OPEN_QTY = 0
        RSTTRXFILE!OPEN_VAL = 0
        RSTTRXFILE!RCPT_QTY = 0
        RSTTRXFILE!RCPT_VAL = 0
        RSTTRXFILE!ISSUE_QTY = 0
        RSTTRXFILE!ISSUE_VAL = 0
        RSTTRXFILE!CLOSE_QTY = 0
        RSTTRXFILE!CLOSE_VAL = 0
        RSTTRXFILE!DAM_QTY = 0
        RSTTRXFILE!DAM_VAL = 0
        RSTTRXFILE!DISC = 0
        RSTTRXFILE.Update
        RSTTRXFILE.MoveNext
    Loop
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
    
    'db.Execute "delete FROM RTRXFILE WHERE VCH_DATE <# " & Format(MDIMAIN.DTFROM.Value, "MM,DD,YYYY") & " # AND BAL_QTY <=0"
    db.Execute "delete FROM REPLCN"
    db.Execute "delete FROM TEMPCN"
    db.Execute "delete FROM TRANSMAST"
    db.Execute "delete FROM TRANSSUB"
    db.Execute "delete FROM TRXEXPENSE"
    db.Execute "delete FROM TRXFILE"
    db.Execute "delete FROM TRXMAST"
    db.Execute "delete FROM TRXSUB"
    db.Execute "delete FROM VANSTOCK"

End Function

Private Sub txtPassword_LostFocus()
    txtPassword.BackColor = vbWhite
End Sub

Private Sub txtyear_Change()
    lblfinyear.Caption = Val(txtyear.Text) + 1
End Sub

Private Sub txtyear_GotFocus()
    txtyear.SelStart = 0
    txtyear.SelLength = Len(txtyear.Text)
    txtyear.BackColor = &H98F3C1
End Sub

Private Sub txtyear_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyReturn, vbKeyTab
            If txtyear.Text = "" Then Exit Sub
            cmdLogin_Click
        Case vbKeyEscape
            txtPassword.SetFocus
    End Select
End Sub

Private Sub txtyear_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"), Asc("["), Asc("]"), Asc("\")
            KeyAscii = 0
        Case vbKey0 To vbKey9, vbKeyLeft, vbKeyRight, vbKeyBack
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtyear_LostFocus()
    txtyear.BackColor = vbWhite
End Sub

Function GetDBNames()

Dim sql As String
Dim recs As New ADODB.Recordset

'        Call opnDB(0)                                        ' External function to Open a MySql connection
        sql = "show databases;"                              ' set sql
        'recs.Open sql, strConnection, adOpenStatic, adLockReadOnly      ' open the recordset. recs is a global. conn is a global.
        recs.Open sql, db, adOpenStatic, adLockReadOnly
        
        'x = 1
        Me.CmbDB.Clear                                 ' this is a listbox on a form somewhere

        Do Until recs.EOF = True                               ' loop until done
            'ReDim Preserve dbname(x)                         ' set a new item in the array
            'dbname(x) = recs.Fields(0)                         ' the first (and only) field returned is the db name
            If Left(recs.Fields(0), 3) = "inv" Then    ' just get the user database names
                Me.CmbDB.AddItem (Mid(recs.Fields(0), 4))       ' add the db name to the list
            End If
            recs.MoveNext                                      ' see if there are any more databases
            'x = x + 1
        Loop
        recs.Close
        Set recs = Nothing
'        Call closdb                                          ' External function to close the recordset and the connection
        'dbfg = 0                                             ' global so I know where I came from. Will have various
                                                             ' values depending on what I'm trying to do.
        
End Function

