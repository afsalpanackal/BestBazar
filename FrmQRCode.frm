VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmQRCode 
   Caption         =   "QRCode"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   ControlBox      =   0   'False
   Icon            =   "FrmQRCode.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   Begin MSDataGridLib.DataGrid grdtmp 
      Height          =   3330
      Left            =   -330
      TabIndex        =   0
      Top             =   2625
      Visible         =   0   'False
      Width           =   14220
      _ExtentX        =   25083
      _ExtentY        =   5874
      _Version        =   393216
      AllowUpdate     =   -1  'True
      AllowArrows     =   -1  'True
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   0
      HeadLines       =   1
      RowHeight       =   20
      RowDividerStyle =   4
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         SizeMode        =   1
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmQRCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If Left(Trim(TXTITEMCODE.Text), 1) = "#" And Len(Trim(TXTITEMCODE.Text)) > 6 Then
        Dim itemcode As String
        Dim itemqty As Double
        
        itemcode = Mid(Trim(TXTITEMCODE.Text), 2, 5)
        itemqty = Val(Mid(Trim(TXTITEMCODE.Text), 7, 5))
        Set grdtmp.DataSource = Nothing
        If PHYFLAG = True Then
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE ITEM_CODE = '" & itemcode & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
        Else
            PHY.Close
            PHY.Open "Select ITEM_CODE, ITEM_NAME, CLOSE_QTY, SALES_PRICE, SALES_TAX, UNIT, P_RETAIL, COM_FLAG, COM_PER, COM_AMT, CRTN_PACK, P_CRTN, P_WS, P_VAN, CHECK_FLAG, CATEGORY, LOOSE_PACK, PACK_TYPE, WARRANTY, WARRANTY_TYPE, MRP, CUST_DISC, CESS_PER, CESS_AMT, P_LWS, UN_BILL, PRICE5, PRICE6, PRICE7 From ITEMMAST  WHERE ITEM_CODE = '" & itemcode & "' AND (ISNULL(DEAD_STOCK) OR DEAD_STOCK <> 'Y')", db, adOpenStatic, adLockReadOnly
            PHYFLAG = False
        End If
        Set grdtmp.DataSource = PHY
        If PHY.RecordCount > 0 Then
            creditbill.TxtCessPer.Text = IIf(IsNull(grdtmp.Columns(22)), "", grdtmp.Columns(22))
            creditbill.TxtCessAmt.Text = IIf(IsNull(grdtmp.Columns(23)), "", grdtmp.Columns(23))
            creditbill.TXTDISC.Text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
            creditbill.LBLUNBILL.Caption = IIf(IsNull(grdtmp.Columns(25)), "N", grdtmp.Columns(25))
            
            creditbill.TxtMRP.Text = IIf(IsNull(grdtmp.Columns(20)), "", grdtmp.Columns(20))
            Select Case creditbill.cmbtype.ListIndex
                Case 1
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(12)), "", Val(grdtmp.Columns(12)))
                Case 2
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(13)), "", Val(grdtmp.Columns(13)))
                Case 3
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(20)), "", Val(grdtmp.Columns(20)))
                    If Val(creditbill.txtretail.Text) = 0 Then
                        creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                        creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    End If
                Case 4
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(26)), "", Val(grdtmp.Columns(26)))
                Case 5
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(27)), "", Val(grdtmp.Columns(27)))
                Case 6
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(28)), "", Val(grdtmp.Columns(28)))
                Case Else
                    creditbill.txtNetrate.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    creditbill.txtretail.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
                    creditbill.TXTRETAILNOTAX.Text = IIf(IsNull(grdtmp.Columns(6)), "", Val(grdtmp.Columns(6)))
            End Select
            creditbill.LblPack.Text = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            creditbill.lblOr_Pack.Caption = IIf(IsNull(grdtmp.Columns(16)) Or Val(grdtmp.Columns(16)) = 0, "1", grdtmp.Columns(16))
            creditbill.TXTDISC.Text = IIf(IsNull(grdtmp.Columns(21)), "", grdtmp.Columns(21))
            'TXTEXPIRY.Text = IIf(isdate(grdtmp.Columns(22)),Format(grdtmp.Columns(22), "MM/YY"),"  /  ")
            creditbill.TXTITEMCODE.Text = grdtmp.Columns(0)


            
            creditbill.item_change = True
            creditbill.TXTPRODUCT.Text = grdtmp.Columns(1)
            creditbill.item_change = False
            creditbill.txtPrintname.Text = grdtmp.Columns(1)
            Select Case PHY!CHECK_FLAG
                Case "M"
                    creditbill.OPTTaxMRP.value = True
                    creditbill.TXTTAX.Text = grdtmp.Columns(4)
                    creditbill.TXTSALETYPE.Text = "2"
                Case "V"
                    creditbill.OPTVAT.value = True
                    creditbill.TXTSALETYPE.Text = "2"
                    creditbill.TXTTAX.Text = grdtmp.Columns(4)
                Case Else
                    creditbill.TXTSALETYPE.Text = "2"
                    creditbill.optnet.value = True
                    creditbill.TXTTAX.Text = "0"
            End Select
            
            'Call CONTINUE
            creditbill.TXTQTY.Text = itemqty
            Call creditbill.TXTQTY_LostFocus
            If MDIMAIN.LblKFCNet.Caption <> "N" Then
                Call creditbill.txtNetrate_LostFocus
                Call creditbill.TXTDISC_LostFocus
                Call creditbill.CMDADD_Click
            Else
                If MDIMAIN.StatusBar.Panels(14).Text <> "Y" Then
                    Call creditbill.TXTRETAILNOTAX_LostFocus
                Else
                    If Val(TxtMRP.Text) <> 0 And Val(TxtMRP.Text) = Val(TXTRETAILNOTAX.Text) And mrpplus = True Then
                        Call creditbill.TXTRETAILNOTAX_LostFocus
                    ElseIf Val(TxtMRP.Text) <> 0 And Val(Round(Val(TxtMRP.Text), 2)) = Val(Round(Val(txtretail.Text), 2)) And mrpplus = False Then
                        Call creditbill.txtNetrate_LostFocus
                    Else
                        Call creditbill.TXTRETAIL_LostFocus
                    End If
                End If
                Call creditbill.TXTDISC_LostFocus
                Call creditbill.CMDADD_Click
            End If
            Exit Sub
            
        End If
    End If
    
End Sub
