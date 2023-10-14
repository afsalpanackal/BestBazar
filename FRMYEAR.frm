VERSION 5.00
Begin VB.Form FRMYEAR 
   BackColor       =   &H00EADFAE&
   BorderStyle     =   0  'None
   Caption         =   "FINANCIAL YEAR"
   ClientHeight    =   1545
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4560
   Icon            =   "FRMYEAR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2205
      TabIndex        =   2
      Top             =   840
      Width           =   1065
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3345
      TabIndex        =   1
      Top             =   840
      Width           =   1035
   End
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
      ForeColor       =   &H00004000&
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1755
      MaxLength       =   4
      TabIndex        =   0
      Top             =   210
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
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   5
      Top             =   255
      Width           =   1635
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
      ForeColor       =   &H00004000&
      Height          =   390
      Left            =   3195
      TabIndex        =   4
      Top             =   210
      Width           =   1170
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00EADFAE&
      Caption         =   "-"
      ForeColor       =   &H00004000&
      Height          =   195
      Left            =   2940
      TabIndex        =   3
      Top             =   225
      Width           =   210
   End
End
Attribute VB_Name = "FRMYEAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub cmdLogin_Click()
    If Len(txtyear.Text) <> 4 Then
        MsgBox "Please enter proper year", vbOKOnly, "Financial Year"
        txtyear.SetFocus
        Exit Sub
    End If
    
'    If Val(txtyear.text) < 2010 Or Val(txtyear.text) > 2030 Then
'        MsgBox "Please enter proper year", vbOKOnly, "Financial Year"
'        txtyear.SetFocus
'        Exit Sub
'    End If
    
    If Val(txtyear.Text) < 2010 Or Val(txtyear.Text) > 2030 Then
        'MsgBox "Please enter proper year", vbOKOnly, "Financial Year"
        MsgBox "Unexpected error occured", vbOKOnly, "Financial Year"
        txtyear.SetFocus
        Exit Sub
    End If


    MDIMAIN.DTFROM.Value = "01/04/" & txtyear.Text
    MDIMAIN.Caption = "EzBiz Inventory - Financial Year " & Trim(txtyear.Text) & " - " & Trim(lblfinyear.Caption)
    
    On Error GoTo eRRhAND
    Dim RSTCOMPANY As ADODB.Recordset
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM COMPINFO WHERE COMP_CODE = '001' AND FIN_YR = '" & Year(MDIMAIN.DTFROM.Value) & "'", db, adOpenStatic, adLockOptimistic, adCmdText
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
        MDIMAIN.LBLLABELNOS.Caption = IIf(IsNull(RSTCOMPANY!BAR_LABELS) Or RSTCOMPANY!BAR_LABELS = 0 Or RSTCOMPANY!BAR_LABELS = "", "1", RSTCOMPANY!BAR_LABELS)
        If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
            MDIMAIN.CmdGST.Caption = "Sales"
            MDIMAIN.CmdRetailBill.Caption = "Sales"
            MDIMAIN.CmdRetailBill2.Caption = "Sales"
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
        If RSTCOMPANY!SHOP_RT = "Y" Then
            MDIMAIN.LBLSHOPRT.Caption = "Y"
        Else
            MDIMAIN.LBLSHOPRT.Caption = "N"
        End If
        If RSTCOMPANY!RST_BILL = "Y" Then
            RstBill_Flag = "Y"
        Else
            RstBill_Flag = "N"
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
        On Error GoTo eRRhAND
    Else
        On Error Resume Next
        Dim RSTCOMPANY2 As ADODB.Recordset
        Set RSTCOMPANY2 = New ADODB.Recordset
        RSTCOMPANY2.Open "Select * From COMPINFO WHERE COMP_CODE = '001' ORDER BY FIN_YR DESC", db, adOpenStatic, adLockReadOnly
        If Not (RSTCOMPANY2.EOF And RSTCOMPANY2.BOF) Then
            RSTCOMPANY.AddNew
            RSTCOMPANY!COMP_CODE = "001"
            RSTCOMPANY!CUST_CODE = RSTCOMPANY2!CUST_CODE
            RSTCOMPANY!SERVER_ADD = RSTCOMPANY2!SERVER_ADD
            RSTCOMPANY!CLOUD_TYPE = RSTCOMPANY2!CLOUD_TYPE
            RSTCOMPANY!MULTI_FLAG = RSTCOMPANY2!MULTI_FLAG
            RSTCOMPANY!REMOTE_FLAG = RSTCOMPANY2!REMOTE_FLAG
            RSTCOMPANY!FIN_YR = Year(MDIMAIN.DTFROM.Value)
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
                RSTCOMPANY!KFCFROM_DATE = RSTCOMPANY2!KFCFROM_DATE
            End If
            If IsDate(RSTCOMPANY2!KFCTO_DATE) Then
                RSTCOMPANY!KFCTO_DATE = RSTCOMPANY2!KFCTO_DATE
            End If
            RSTCOMPANY!item_repeat = RSTCOMPANY2!item_repeat
            RSTCOMPANY!VS_FLAG = RSTCOMPANY2!VS_FLAG
            RSTCOMPANY!LMT_FLAG = RSTCOMPANY2!LMT_FLAG
            RSTCOMPANY!MINUS_BILL = RSTCOMPANY2!MINUS_BILL
            RSTCOMPANY!BAR_TEMPLATE = RSTCOMPANY2!BAR_TEMPLATE
            RSTCOMPANY!BAR_FORMAT = RSTCOMPANY2!BAR_FORMAT
            RSTCOMPANY!MOB_WARN_FLAG = RSTCOMPANY2!MOB_WARN_FLAG
            RSTCOMPANY!CLCODE = RSTCOMPANY2!CLCODE
            RSTCOMPANY!PCODE = RSTCOMPANY2!PCODE
            RSTCOMPANY!SCHEME_OPT = RSTCOMPANY2!SCHEME_OPT
            RSTCOMPANY.Update
            
            On Error GoTo eRRhAND
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
            On Error GoTo eRRhAND
        End If
        RSTCOMPANY2.Close
        Set RSTCOMPANY2 = Nothing
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Unload Me
    Exit Sub
eRRhAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub Form_Load()
    MDIMAIN.Enabled = False
    txtyear.Text = MDIMAIN.DTFROM.Year
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Top = 4000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MDIMAIN.Enabled = True
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
            Unload Me
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

