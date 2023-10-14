VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{55473EAC-7715-4257-B5EF-6E14EBD6A5DD}#1.0#0"; "vbalProgBar6.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.MDIForm MDIMAIN 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Ez..Biz INVENTORY"
   ClientHeight    =   10905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   19110
   Icon            =   "MDIMAIN.frx":0000
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   Picture         =   "MDIMAIN.frx":08CA
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      BackColor       =   &H00FFFFFF&
      DragMode        =   1  'Automatic
      DrawStyle       =   1  'Dash
      Height          =   9675
      Left            =   18195
      ScaleHeight     =   9615
      ScaleWidth      =   855
      TabIndex        =   44
      Top             =   780
      Width           =   915
      Begin MSForms.CommandButton CmdDeposit 
         Height          =   465
         Left            =   -15
         TabIndex        =   52
         Top             =   6915
         Width           =   870
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Deposit"
         PicturePosition =   327683
         Size            =   "1535;820"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton2 
         Height          =   465
         Left            =   -15
         TabIndex        =   25
         Top             =   4455
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Contra"
         PicturePosition =   327683
         Size            =   "1773;820"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdorder 
         Height          =   540
         Left            =   -15
         TabIndex        =   16
         Top             =   -60
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Sales Reg (Ctrl+R)"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdExpiry 
         Height          =   450
         Left            =   -15
         TabIndex        =   33
         Top             =   8910
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Expiry"
         PicturePosition =   327683
         Size            =   "1773;794"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdSupMast 
         Height          =   540
         Left            =   -15
         TabIndex        =   29
         Top             =   6375
         Width           =   945
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Supp Mast(Ctrl+N)"
         PicturePosition =   327683
         Size            =   "1667;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdCust 
         Height          =   540
         Left            =   -15
         TabIndex        =   28
         Top             =   5850
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Cust Mast(Ctrl+M)"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdItemMast 
         Height          =   540
         Left            =   -15
         TabIndex        =   27
         Top             =   5340
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Item Mast(Ctrl+C)"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdPayment 
         Height          =   540
         Left            =   -15
         TabIndex        =   17
         Top             =   435
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Payment F9"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdReceipt 
         Height          =   540
         Left            =   -15
         TabIndex        =   18
         Top             =   945
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Receipt   F8"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdExp 
         Height          =   540
         Left            =   -15
         TabIndex        =   19
         Top             =   1455
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Off Exp F11"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdStaff 
         Height          =   570
         Left            =   -15
         TabIndex        =   20
         Top             =   1935
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Salary Register"
         PicturePosition =   327683
         Size            =   "1773;1005"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdLend 
         Height          =   540
         Left            =   -15
         TabIndex        =   21
         Top             =   2475
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Money Lend"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdBook 
         Height          =   465
         Left            =   -15
         TabIndex        =   26
         Top             =   4905
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Cash Book"
         PicturePosition =   327683
         Size            =   "1773;820"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmDIncome 
         Height          =   450
         Left            =   -15
         TabIndex        =   22
         Top             =   2985
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Off Income"
         PicturePosition =   327683
         Size            =   "1773;794"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdJournal 
         Height          =   540
         Left            =   -15
         TabIndex        =   23
         Top             =   3420
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Dr/ Cr Note Entry"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   135
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdSearch 
         Height          =   540
         Left            =   -15
         TabIndex        =   31
         Top             =   7890
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Search (Ctrl+S)"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmDStkMvmnt 
         Height          =   540
         Left            =   -15
         TabIndex        =   30
         Top             =   7380
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Stock Mvmnt"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CmdQutn 
         Height          =   555
         Left            =   -15
         TabIndex        =   32
         Top             =   8385
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Qutns(Ctl+Q)"
         PicturePosition =   327683
         Size            =   "1773;979"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton CommandButton3 
         Height          =   540
         Left            =   -15
         TabIndex        =   24
         Top             =   3945
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Journal (Ctrl +J)"
         PicturePosition =   327683
         Size            =   "1773;952"
         FontName        =   "Tahoma"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox PCTMENU 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   0
      ScaleHeight     =   780
      ScaleWidth      =   19110
      TabIndex        =   35
      Top             =   0
      Width           =   19110
      Begin VB.TextBox TxtDUP 
         Appearance      =   0  'Flat
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   15990
         MaxLength       =   4
         PasswordChar    =   "@"
         TabIndex        =   87
         Top             =   450
         Visible         =   0   'False
         Width           =   1050
      End
      Begin VB.CommandButton cmdsendsms 
         Caption         =   "Send SMS"
         Height          =   375
         Left            =   14505
         TabIndex        =   76
         Top             =   420
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   570
         Left            =   13725
         TabIndex        =   43
         Top             =   795
         Visible         =   0   'False
         Width           =   585
      End
      Begin vbalProgBarLib6.vbalProgressBar vbalProgressBar1 
         Height          =   735
         Left            =   15975
         TabIndex        =   36
         Tag             =   "5"
         Top             =   30
         Visible         =   0   'False
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1296
         Picture         =   "MDIMAIN.frx":C117
         Appearance      =   0
         BarPicture      =   "MDIMAIN.frx":C133
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DTFROM 
         Height          =   600
         Left            =   120
         TabIndex        =   42
         Top             =   780
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   1058
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         Format          =   147652609
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTKFCSTART 
         Height          =   300
         Left            =   0
         TabIndex        =   61
         Top             =   810
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         Format          =   147652609
         CurrentDate     =   40498
      End
      Begin MSComCtl2.DTPicker DTKFCEND 
         Height          =   300
         Left            =   1650
         TabIndex        =   62
         Top             =   930
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         CalendarForeColor=   0
         CalendarTitleForeColor=   16576
         CalendarTrailingForeColor=   255
         CheckBox        =   -1  'True
         Format          =   147652609
         CurrentDate     =   40498
      End
      Begin VB.Label LBLTRCopy 
         Height          =   120
         Left            =   0
         TabIndex        =   86
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label LBLLABELNOS 
         Height          =   150
         Left            =   0
         TabIndex        =   85
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblec 
         Height          =   120
         Left            =   0
         TabIndex        =   84
         Top             =   0
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSForms.CommandButton CmdRetailBill2 
         Height          =   765
         Left            =   810
         TabIndex        =   1
         Top             =   0
         Width           =   885
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Sales B2C"
         PicturePosition =   327683
         Size            =   "1561;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label lblPerPurchase 
         Height          =   135
         Left            =   0
         TabIndex        =   83
         Top             =   765
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblmobwarn 
         Height          =   135
         Left            =   0
         TabIndex        =   82
         Top             =   765
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblPriceSplit 
         Height          =   120
         Left            =   0
         TabIndex        =   81
         Top             =   810
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblcategory 
         Height          =   135
         Left            =   0
         TabIndex        =   80
         Top             =   780
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lblsalary 
         Height          =   165
         Left            =   1230
         TabIndex        =   79
         Top             =   555
         Visible         =   0   'False
         Width           =   150
      End
      Begin MSForms.CommandButton CmdVsale 
         Height          =   765
         Left            =   13815
         TabIndex        =   78
         Top             =   0
         Width           =   1035
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Branch Sale"
         PicturePosition =   327683
         Size            =   "1826;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label LBLITMWRN 
         Height          =   135
         Left            =   2865
         TabIndex        =   77
         Top             =   600
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label lbldmpmini 
         Height          =   90
         Left            =   5010
         TabIndex        =   75
         Top             =   615
         Visible         =   0   'False
         Width           =   165
      End
      Begin MSForms.CommandButton CmdSmry 
         Height          =   780
         Left            =   3780
         TabIndex        =   4
         Top             =   -15
         Width           =   1005
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Detailed Stock"
         PicturePosition =   327683
         Size            =   "1773;1376"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin VB.Label LBLSTATENAME 
         Height          =   165
         Left            =   0
         TabIndex        =   74
         Top             =   855
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label LBLSTATE 
         Height          =   165
         Left            =   3570
         TabIndex        =   73
         Top             =   615
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblExpEnable 
         Height          =   150
         Left            =   2475
         TabIndex        =   72
         Top             =   645
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblRemoveUbill 
         Height          =   120
         Left            =   1830
         TabIndex        =   71
         Top             =   585
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label lblprnall 
         Height          =   150
         Left            =   1470
         TabIndex        =   70
         Top             =   540
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Label LBLAMC 
         Height          =   300
         Left            =   0
         TabIndex        =   69
         Top             =   855
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label LBLSHOPRT 
         Height          =   105
         Left            =   0
         TabIndex        =   68
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblitemrepeat 
         Height          =   105
         Left            =   0
         TabIndex        =   67
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label LBLSPACE 
         Height          =   105
         Left            =   0
         TabIndex        =   66
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label LBLMRPPLUS 
         Height          =   105
         Left            =   0
         TabIndex        =   65
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label LBLHSNSUM 
         Height          =   105
         Left            =   0
         TabIndex        =   64
         Top             =   0
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label LblKFCNet 
         Height          =   300
         Left            =   0
         TabIndex        =   63
         Top             =   765
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblkfc 
         Height          =   195
         Left            =   0
         TabIndex        =   60
         Top             =   840
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblmrp 
         Height          =   195
         Left            =   0
         TabIndex        =   59
         Top             =   840
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblprint 
         Height          =   105
         Left            =   750
         TabIndex        =   58
         Top             =   930
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Label lblub 
         Height          =   195
         Left            =   0
         TabIndex        =   57
         Top             =   855
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label lblvp 
         Height          =   195
         Left            =   0
         TabIndex        =   56
         Top             =   945
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label LBLWS 
         Height          =   195
         Left            =   0
         TabIndex        =   55
         Top             =   975
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label LBLRT 
         Height          =   195
         Left            =   13635
         TabIndex        =   54
         Top             =   705
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.Label LBLGSTWRN 
         Height          =   300
         Left            =   0
         TabIndex        =   53
         Top             =   870
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblnostock 
         Height          =   195
         Left            =   3255
         TabIndex        =   51
         Top             =   795
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label barcode_profile 
         Height          =   135
         Left            =   13890
         TabIndex        =   50
         Top             =   615
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label LBLTHPREVIEW 
         Height          =   210
         Left            =   14160
         TabIndex        =   49
         Top             =   300
         Visible         =   0   'False
         Width           =   150
      End
      Begin VB.Label LBLDMPTHERMAL 
         Height          =   210
         Left            =   14130
         TabIndex        =   48
         Top             =   165
         Visible         =   0   'False
         Width           =   180
      End
      Begin VB.Label lblgst 
         Height          =   300
         Left            =   13980
         TabIndex        =   47
         Top             =   495
         Visible         =   0   'False
         Width           =   360
      End
      Begin VB.Label lblform62 
         Height          =   315
         Left            =   0
         TabIndex        =   46
         Top             =   825
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label LBLTAXWARN 
         Height          =   315
         Left            =   14235
         TabIndex        =   45
         Top             =   135
         Visible         =   0   'False
         Width           =   210
      End
      Begin MSForms.CommandButton CmdGST 
         Height          =   765
         Left            =   4785
         TabIndex        =   5
         Top             =   0
         Width           =   1035
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Sales B2B"
         PicturePosition =   327683
         Size            =   "1826;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton4 
         Height          =   765
         Left            =   18345
         TabIndex        =   14
         Top             =   0
         Width           =   1020
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Process"
         PicturePosition =   327683
         Size            =   "1799;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdDelivery 
         Height          =   765
         Left            =   12795
         TabIndex        =   12
         Top             =   0
         Width           =   1020
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Delivery"
         PicturePosition =   327683
         Size            =   "1799;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDTRANSFER 
         Height          =   765
         Left            =   11805
         TabIndex        =   11
         Top             =   0
         Width           =   990
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Stock Transfer"
         PicturePosition =   327683
         Size            =   "1746;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMD62 
         Height          =   765
         Left            =   10770
         TabIndex        =   10
         Top             =   0
         Width           =   1035
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Local Purchase"
         PicturePosition =   327683
         Size            =   "1826;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdProduction 
         Height          =   765
         Left            =   9540
         TabIndex        =   9
         Top             =   0
         Width           =   1230
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Production"
         PicturePosition =   327683
         Size            =   "2170;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CommandButton1 
         Height          =   765
         Left            =   19380
         TabIndex        =   15
         Top             =   0
         Width           =   855
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Refresh Stock"
         PicturePosition =   327683
         Size            =   "1508;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDDUPPURCHASE 
         Height          =   690
         Left            =   13395
         TabIndex        =   41
         Top             =   780
         Visible         =   0   'False
         Width           =   855
         Caption         =   "DUPLICATE BILL"
         PicturePosition =   327683
         Size            =   "1508;1217"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton cmdPOrde1 
         Height          =   675
         Left            =   18690
         TabIndex        =   40
         Top             =   30
         Visible         =   0   'False
         Width           =   540
         ForeColor       =   16777215
         BackColor       =   255
         Caption         =   "ORDER"
         PicturePosition =   327683
         Size            =   "952;1191"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdduplicate 
         Height          =   675
         Left            =   18690
         TabIndex        =   39
         Top             =   15
         Visible         =   0   'False
         Width           =   1185
         Caption         =   "DUPLICATE BILL"
         PicturePosition =   327683
         Size            =   "2090;1191"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
      Begin MSForms.CommandButton STKADJUST 
         Height          =   765
         Left            =   2715
         TabIndex        =   3
         Top             =   0
         Width           =   1065
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Stock Summary"
         PicturePosition =   327683
         Size            =   "1879;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdpurchase 
         Height          =   765
         Left            =   1695
         TabIndex        =   2
         Top             =   0
         Width           =   1020
         ForeColor       =   0
         BackColor       =   2669040
         Caption         =   "Purchase"
         PicturePosition =   327683
         Size            =   "1799;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDSALERETURN 
         Height          =   765
         Left            =   6870
         TabIndex        =   7
         Top             =   0
         Width           =   1005
         ForeColor       =   16777215
         BackColor       =   8421376
         VariousPropertyBits=   8388635
         Caption         =   "Sales Return"
         PicturePosition =   327683
         Size            =   "1773;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDPRETURN 
         Height          =   765
         Left            =   5820
         TabIndex        =   6
         Top             =   0
         Width           =   1050
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Purchase Return"
         PicturePosition =   327683
         Size            =   "1852;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDBILLRETAIL 
         Height          =   765
         Left            =   7875
         TabIndex        =   8
         Top             =   0
         Width           =   1650
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Service Bills / Delivery Chellan/ Exemp. Sales"
         PicturePosition =   327683
         Size            =   "2910;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton cmdtest 
         Height          =   405
         Left            =   14490
         TabIndex        =   38
         Top             =   450
         Visible         =   0   'False
         Width           =   990
         ForeColor       =   16777215
         BackColor       =   128
         VariousPropertyBits=   8388635
         Caption         =   "test"
         PicturePosition =   327683
         Size            =   "1746;714"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdWO 
         Height          =   390
         Left            =   18690
         TabIndex        =   37
         Top             =   0
         Visible         =   0   'False
         Width           =   945
         ForeColor       =   16777215
         BackColor       =   16777215
         PicturePosition =   327683
         Size            =   "1667;688"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdRetailBill 
         Height          =   765
         Left            =   0
         TabIndex        =   0
         Top             =   0
         Width           =   810
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "Sales B2C"
         PicturePosition =   327683
         Size            =   "1429;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   180
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CMDEXIT 
         Height          =   765
         Left            =   14850
         TabIndex        =   13
         Top             =   0
         Width           =   1110
         ForeColor       =   0
         BackColor       =   2669040
         VariousPropertyBits=   8388635
         Caption         =   "EXIT"
         PicturePosition =   327683
         Size            =   "1958;1349"
         FontName        =   "Tahoma"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
   Begin ComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      TabIndex        =   34
      Top             =   10455
      Width           =   19110
      _ExtentX        =   33708
      _ExtentY        =   794
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   16
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   2
            Object.Width           =   6165
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Width           =   8890
            MinWidth        =   8890
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel8 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel9 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel10 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel11 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel12 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel13 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel14 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Visible         =   0   'False
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel15 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel16 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu MNUENTRY 
      Caption         =   "ENTRIES"
      Begin VB.Menu mnusale 
         Caption         =   "Sales"
         Begin VB.Menu mnuwsale1 
            Caption         =   "GST B2B"
            Shortcut        =   +{F1}
         End
         Begin VB.Menu mnuwsale 
            Caption         =   "GST B2B"
            Shortcut        =   +{F2}
         End
         Begin VB.Menu MNUGSTBTBU 
            Caption         =   "GSTB2B (U)"
            Shortcut        =   ^W
         End
         Begin VB.Menu MNUGSTSERVICE 
            Caption         =   "GSTSERVICE (U)"
            Shortcut        =   ^Y
         End
         Begin VB.Menu mnub2c1 
            Caption         =   "GST B2C -1"
            Shortcut        =   ^{F1}
         End
         Begin VB.Menu mnu2 
            Caption         =   "GST B2C - 2"
            Shortcut        =   ^{F2}
         End
         Begin VB.Menu mnuretail 
            Caption         =   "GST B2C - 3"
            Shortcut        =   ^{F3}
         End
         Begin VB.Menu mnutblewise 
            Caption         =   "Table Wise"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu mnupurch 
         Caption         =   "Purchase"
         Begin VB.Menu mnupurchase 
            Caption         =   "Purchase Bill"
            Shortcut        =   ^P
         End
         Begin VB.Menu mnulocal 
            Caption         =   "Local Purchase"
         End
         Begin VB.Menu MunUPurchase 
            Caption         =   "Purchase (U)"
            Shortcut        =   ^X
         End
         Begin VB.Menu mnuAstPurchase 
            Caption         =   "Assets Purchase"
         End
         Begin VB.Menu mnuExpPurchase 
            Caption         =   "Expense Entry (Input Tax to be Credited)"
         End
      End
      Begin VB.Menu mnustock 
         Caption         =   "Stock"
         Begin VB.Menu MNUPRICE 
            Caption         =   "Price Analysis"
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuloc 
            Caption         =   "Search Item Location"
            Shortcut        =   ^L
         End
         Begin VB.Menu mnustkcr 
            Caption         =   "Stock Correction"
            Shortcut        =   {F1}
         End
         Begin VB.Menu mnustksum 
            Caption         =   "Stock Summary"
         End
         Begin VB.Menu mnustk 
            Caption         =   "Detailed Stock Analysis"
         End
         Begin VB.Menu MNUOPSTK 
            Caption         =   "Stock Adjustment Entry"
         End
      End
      Begin VB.Menu mnuexpret 
         Caption         =   "Expiry Return by Customers"
      End
      Begin VB.Menu mnuhead 
         Caption         =   "Head Creation"
         Begin VB.Menu mnuitemmast 
            Caption         =   "Item Master"
            Shortcut        =   ^C
         End
         Begin VB.Menu MNUSUPPLIER 
            Caption         =   "Supplier Master"
            Shortcut        =   ^N
         End
         Begin VB.Menu MNUCUST 
            Caption         =   "Customer Master"
            Shortcut        =   ^M
         End
         Begin VB.Menu MNUEMPLOYEE 
            Caption         =   "Employee Master"
         End
         Begin VB.Menu MNULEND 
            Caption         =   "Lender Master"
         End
         Begin VB.Menu mnudeposit 
            Caption         =   "Deposit / Savings Master"
         End
         Begin VB.Menu MNUEXPLDGR 
            Caption         =   "Expense Master"
         End
         Begin VB.Menu MnuIncomeMast 
            Caption         =   "Income Master"
         End
         Begin VB.Menu MNUASTMASTR 
            Caption         =   "Assets Master"
         End
         Begin VB.Menu MNUAGNT 
            Caption         =   "Salesman / Agent Master"
         End
         Begin VB.Menu mnutable 
            Caption         =   "Table Creation"
         End
         Begin VB.Menu MNUBANKMSTR 
            Caption         =   "Bank Master"
         End
         Begin VB.Menu mnubrmaster 
            Caption         =   "Branch Master"
         End
         Begin VB.Menu mnuvehicle 
            Caption         =   "Vehicle Master"
            Shortcut        =   ^H
         End
         Begin VB.Menu mnucat 
            Caption         =   "Category Master"
         End
      End
      Begin VB.Menu mnu3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPlu 
         Caption         =   "PLU MASTER"
         Begin VB.Menu mnuPluItems 
            Caption         =   "Item Creation with PLU Code"
            Shortcut        =   ^U
         End
         Begin VB.Menu mnuexportMc 
            Caption         =   "Export Items to the Machine"
         End
      End
      Begin VB.Menu mnu1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuvoucher 
         Caption         =   "Voucher Entry"
         Begin VB.Menu MnuExpenseSt 
            Caption         =   "Expense to Staff"
            Shortcut        =   {F12}
         End
         Begin VB.Menu MNUEXPENTRY 
            Caption         =   "Office Expense Entry"
            Shortcut        =   {F11}
         End
         Begin VB.Menu mnuFxdAssets 
            Caption         =   "Fixed Assets Entry "
         End
         Begin VB.Menu mnujournal 
            Caption         =   "Journal"
            Shortcut        =   ^J
         End
         Begin VB.Menu mnuincom 
            Caption         =   "Office Income"
         End
         Begin VB.Menu PGAYR 
            Caption         =   "Payment Entry"
            Shortcut        =   {F9}
         End
         Begin VB.Menu mnurcptregstr 
            Caption         =   "Receipt Entry"
            Shortcut        =   {F8}
         End
         Begin VB.Menu mnucontra 
            Caption         =   "Contra"
         End
         Begin VB.Menu MNYLND 
            Caption         =   "Money Lend"
         End
         Begin VB.Menu mnudepositentry 
            Caption         =   "Deposit"
         End
      End
      Begin VB.Menu MnuMinQty 
         Caption         =   "Set Re-Order Qty"
      End
      Begin VB.Menu mnuOPCash 
         Caption         =   "Opening Cash"
      End
      Begin VB.Menu mnuFormatSep 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSer_Stk 
         Caption         =   "Issue Items to Service"
      End
      Begin VB.Menu MNUQTN 
         Caption         =   "Quotations"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnotkorder 
         Caption         =   "Take Order From Customers"
      End
      Begin VB.Menu MNUPO 
         Caption         =   "Purchase Order"
      End
      Begin VB.Menu mnuproduction 
         Caption         =   "General Production"
      End
      Begin VB.Menu MNUOP 
         Caption         =   "Opening Stock Entry"
      End
      Begin VB.Menu mnuimport 
         Caption         =   "Import Purchase / Stock Transfer"
      End
      Begin VB.Menu mnustktransfer 
         Caption         =   "Stock Transfer from HO / Other Branch"
      End
      Begin VB.Menu MNUDAMAGE 
         Caption         =   "Damage Goods Entry"
      End
      Begin VB.Menu mnudmgbr 
         Caption         =   "Damage Goods Entry(frm Branch)"
      End
      Begin VB.Menu mnudamge 
         Caption         =   "Damage Goods from Customers (Credit Note)"
      End
      Begin VB.Menu MNUGIFT 
         Caption         =   "Sample Goods Entry"
      End
      Begin VB.Menu MNUBANKBOOK 
         Caption         =   "BANK BOOK"
      End
      Begin VB.Menu mnulINE 
         Caption         =   "-"
      End
      Begin VB.Menu MNUFMLA 
         Caption         =   "PRODUCTION FORMULA"
      End
      Begin VB.Menu MNUPROCESS 
         Caption         =   "PROCESS FORMULA"
      End
      Begin VB.Menu mnuezbill 
         Caption         =   "EzBill"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu MNUREPORT 
      Caption         =   "VIEW"
      Begin VB.Menu mnureg 
         Caption         =   "Registers"
         Begin VB.Menu SR 
            Caption         =   "Sales Register"
            Shortcut        =   ^R
         End
         Begin VB.Menu mnuBrsalereg 
            Caption         =   "Branch Sale Register"
         End
         Begin VB.Menu PR 
            Caption         =   "Purchase Register"
         End
         Begin VB.Menu Mnuexchange 
            Caption         =   "Exchange Register"
         End
         Begin VB.Menu mnudncnreg 
            Caption         =   "Credit / Debit Note Register"
         End
         Begin VB.Menu mnuassets 
            Caption         =   "Assets Purchase Register"
         End
         Begin VB.Menu mnuExpRegtax 
            Caption         =   "Expense Register (Input Tax)"
         End
         Begin VB.Menu MNUPROREG 
            Caption         =   "Production Register"
         End
         Begin VB.Menu MNUQTNVIEW 
            Caption         =   "Quotation Register"
            Shortcut        =   ^B
         End
         Begin VB.Menu mnuCounter 
            Caption         =   "Counter Register"
            Shortcut        =   ^E
         End
         Begin VB.Menu MnuSerStkReg 
            Caption         =   "Service Stock register"
         End
         Begin VB.Menu mnurcptreg 
            Caption         =   "Receipt Register"
         End
         Begin VB.Menu mnuExpReg 
            Caption         =   "Office Expense Register"
         End
         Begin VB.Menu MnuExpStaffReg 
            Caption         =   "Staff Expense Register"
         End
         Begin VB.Menu mnuFAReg 
            Caption         =   "Fixed Assets Register"
         End
         Begin VB.Menu mnuincome 
            Caption         =   "Income Register"
         End
         Begin VB.Menu MNUDEL_REG 
            Caption         =   "Delivery Register"
         End
         Begin VB.Menu MNUS_RETURNREG 
            Caption         =   "Sales Return Register"
         End
         Begin VB.Menu MNUP_RETURNREG 
            Caption         =   "Purchase Return Register"
         End
         Begin VB.Menu mnudelitems 
            Caption         =   "Deleted Items Regsiter"
         End
         Begin VB.Menu MNUDAMREG 
            Caption         =   "Damaged Goods Register"
         End
         Begin VB.Menu mnubr 
            Caption         =   "Damaged Goods Register Branch Sale"
         End
         Begin VB.Menu MNUGIFTREG 
            Caption         =   "Sample Goods Register"
         End
         Begin VB.Menu MNUCOMSN 
            Caption         =   "Commission Register"
         End
         Begin VB.Menu mnupromotion 
            Caption         =   "Promotion Register"
         End
         Begin VB.Menu MNUSERVICES 
            Caption         =   "Services Register"
         End
         Begin VB.Menu MNUFREE 
            Caption         =   "Free Items Register"
         End
      End
      Begin VB.Menu mnusalesrep 
         Caption         =   "Sales && Purchase Report for E-Filing"
      End
      Begin VB.Menu MNUTRNX 
         Caption         =   "TRANSACTIONS"
      End
      Begin VB.Menu mnuorder 
         Caption         =   "Place Order"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnurcptdue 
         Caption         =   "Receipt Dues"
      End
      Begin VB.Menu mnupymntdue 
         Caption         =   "Payment Dues"
      End
      Begin VB.Menu mnustkrpt 
         Caption         =   "Stock Report"
         Begin VB.Menu STKMVMNT 
            Caption         =   "Stock Movement"
         End
         Begin VB.Menu mnuinout 
            Caption         =   "Inward / Outward Details of All Items"
         End
         Begin VB.Menu n 
            Caption         =   "-"
         End
         Begin VB.Menu mnustkmv 
            Caption         =   "Stock Movement (Branch Sale)"
         End
         Begin VB.Menu mnuinoutbr 
            Caption         =   "Inward / Outward Details of All Branch Sale Items "
         End
         Begin VB.Menu k 
            Caption         =   "-"
         End
         Begin VB.Menu mnuarea 
            Caption         =   "Stock Movement Area Wise"
         End
         Begin VB.Menu mnuStkAnalysis 
            Caption         =   "Stock Analysis"
         End
         Begin VB.Menu mnulabel 
            Caption         =   "Label Printing for Self Items"
            Shortcut        =   ^Z
         End
         Begin VB.Menu MNUPRICELIST 
            Caption         =   "Price List (Retail)"
            Shortcut        =   {F4}
         End
         Begin VB.Menu mnuplist 
            Caption         =   "Price List"
            Shortcut        =   ^G
         End
         Begin VB.Menu mnurordr 
            Caption         =   "Re Order Stock List"
         End
         Begin VB.Menu mnuzero 
            Caption         =   "Zero Stock Items"
         End
         Begin VB.Menu mnuDead 
            Caption         =   "Dead Stock Items"
         End
      End
      Begin VB.Menu MNUCUSTLIST 
         Caption         =   "Customer && Supplier List"
      End
      Begin VB.Menu mnumnth 
         Caption         =   "Customer && Supplier Monthly Analysis"
      End
      Begin VB.Menu MNUPRINTMIX 
         Caption         =   "Print Mixture Formula"
      End
      Begin VB.Menu mnuaccbk 
         Caption         =   "Books of Accounts"
         Begin VB.Menu mnuaccounts 
            Caption         =   "Accounts Summary"
            Shortcut        =   ^A
         End
         Begin VB.Menu mnuabstract 
            Caption         =   "Ledger Abstract"
         End
         Begin VB.Menu MNUCASHBOOK 
            Caption         =   "Cash Book"
         End
         Begin VB.Menu mnutrial 
            Caption         =   "Trial  Balance"
         End
      End
      Begin VB.Menu MNUROUTE 
         Caption         =   "Route Layout"
      End
      Begin VB.Menu mnuReminder 
         Caption         =   "Receipt Dues & Entries"
      End
      Begin VB.Menu mnuAMC 
         Caption         =   "AMC Reminder"
      End
      Begin VB.Menu mnuremind 
         Caption         =   "Birthday / Marriage Reminder"
      End
   End
   Begin VB.Menu mnutest 
      Caption         =   "Test"
      Visible         =   0   'False
   End
   Begin VB.Menu mnugud_rep 
      Caption         =   "ITEMS UNDER WARRANTY"
      Visible         =   0   'False
      Begin VB.Menu MNURCVGOODS 
         Caption         =   "Receive Items from Customer"
      End
      Begin VB.Menu mnusentgoods 
         Caption         =   "Sent Items to the Supplier"
      End
      Begin VB.Menu mnurcvdist 
         Caption         =   "Receive Items from Distributors"
      End
      Begin VB.Menu mnureturn 
         Caption         =   "Return Items to Customers"
      End
   End
   Begin VB.Menu MNUTOOLS 
      Caption         =   "TOOLS"
      Begin VB.Menu mnuexport 
         Caption         =   "EXPORT ALL ITEMS"
         Visible         =   0   'False
      End
      Begin VB.Menu MNUSHOPINFO 
         Caption         =   "EDIT SHOP INFORMATION"
      End
      Begin VB.Menu MNUYEAR 
         Caption         =   "CHANGE FINANCIAL YEAR"
      End
      Begin VB.Menu MNUUSER 
         Caption         =   "EDIT USER PASSWORD"
      End
      Begin VB.Menu MNUBACK 
         Caption         =   "CREATE BACKUP"
      End
      Begin VB.Menu mnufix 
         Caption         =   "FIX TABLE ERROR"
      End
      Begin VB.Menu mnusync 
         Caption         =   "Sync with Cloud"
      End
      Begin VB.Menu mnumerge 
         Caption         =   "ITEM MERGE"
      End
      Begin VB.Menu mnumergemulti 
         Caption         =   "ITEM MERGE (Multiple Items)"
      End
      Begin VB.Menu mn1 
         Caption         =   "-"
      End
      Begin VB.Menu MNUCOST 
         Caption         =   "REFRESH COST &&  PRICE"
      End
      Begin VB.Menu MNUIPCOST 
         Caption         =   "UPDATE OPSTOCK COST"
      End
      Begin VB.Menu MNUREFRESH 
         Caption         =   "REFRESH STOCK"
      End
      Begin VB.Menu MNULITEMLST 
         Caption         =   "RESTORE ITEM LIST"
      End
      Begin VB.Menu MNURESTRECUST 
         Caption         =   "RESTORE CUSTOMER MASTER"
      End
      Begin VB.Menu MNUPRICEUPDATE 
         Caption         =   "UPDATE PRICE LIST"
      End
      Begin VB.Menu RSTBILLS 
         Caption         =   "RESTORE CORRUPT BILLS"
      End
      Begin VB.Menu mnudelzero 
         Caption         =   "DELETE ZERO STOCK ITEMS WITH NIL TRANSACTIONS"
      End
      Begin VB.Menu mn2 
         Caption         =   "-"
      End
      Begin VB.Menu mnucalc 
         Caption         =   "Calculator"
         Shortcut        =   ^K
      End
      Begin VB.Menu mnuabt 
         Caption         =   "About"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Shortcut        =   ^{F4}
      End
   End
End
Attribute VB_Name = "MDIMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim CLOSEALL As Integer
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Const DRIVE_REMOVABLE = 2

Private Sub CMD62_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    If lblform62.Caption = "Y" Then
        frm62.Show
        frm62.SetFocus
    Else
        frmLL.Show
        frmLL.SetFocus
    End If
End Sub

Private Sub CMD62_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CMD62_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CMDBILLRETAIL_Click()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        FRMSERVICEC.Show
        FRMSERVICEC.SetFocus
    Else
        If frmLogin.rs!Level = "5" Then
            Exit Sub
        Else
            FRMSALESSV.Show
            FRMSALESSV.SetFocus
        End If
    End If
                
    
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
End Sub

Private Sub CMDBILLRETAIL_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CMDBILLRETAIL_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdBook_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdCust_Click()
    frmcustmast1.Show
    frmcustmast1.SetFocus
End Sub

Private Sub CmdCust_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdDelivery_Click()
    FRMDELIVERY.Show
    FRMDELIVERY.SetFocus
End Sub

Private Sub CmdDelivery_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdDelivery_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdDeposit_Click()
    FRMDPReg.Show
    FRMDPReg.SetFocus
End Sub

Private Sub cmdduplicate_Click()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    If bill_type_flag = True Then
        FRMPETTY.Show
        FRMPETTY.SetFocus
    Else
        FRMPETTY_TYPE.Show
        FRMPETTY_TYPE.SetFocus
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description, , "EzBiz"
End Sub

Private Sub CMDEXIT_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdExp_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdExpiry_Click()
    FRMEXP.Show
    FRMEXP.SetFocus
End Sub

Private Sub CmdExpiry_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdGST_Click()

'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS3.Show
            FRMPOS3.SetFocus
        Else
            FRMSALES2.Show
            FRMSALES2.SetFocus
        End If
    Else
        If IsFormLoaded(FRMGST) <> True Then
            FRMGST.Show
            FRMGST.SetFocus
        ElseIf IsFormLoaded(FRMGST1) <> True Then
            FRMGST1.Show
            FRMGST1.SetFocus
        ElseIf IsFormLoaded(FRMGST) = True And IsFormLoaded(FRMGST1) = True Then
            FRMGST.Show
            FRMGST.SetFocus
        End If
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
End Sub

Private Sub CmdGST_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdGST_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmDIncome_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdItemMast_Click()
    frmitemmaster.Show
    frmitemmaster.SetFocus
End Sub

Private Sub CmdItemMast_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdJournal_Click()
    FRMDrEntry.Show
    FRMDrEntry.SetFocus
End Sub

Private Sub CmdJournal_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdLend_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub cmdorder_Click()
    FRMBILLPRINT.Show
    FRMBILLPRINT.SetFocus
End Sub

Private Sub CmdPayment_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CMDPRETURN_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdProduction_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdProduction_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub cmdpurchase_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdQutn_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdReceipt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdRetailBill_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
    
End Sub

Private Sub CmdRetailBill2_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
    
End Sub


Private Sub CMDSALERETURN_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CmdSearch_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub cmdsendsms_Click()
    
    Screen.MousePointer = vbHourglass
    Call SendJSONPOSTRequest
'    If IsConnected = False Then
'        Screen.MousePointer = vbNormal
'        MsgBox "You need an internet Connection for translation.", vbOKOnly, "EzBiz"
'        Exit Sub
'    End If
'
'    Dim http As Object
'    Dim url As String
'    Set http = CreateObject("WinHttp.WinHttprequest.5.1")
'    url = "http://192.168.29.182:5500/getorder/"
'    http.Open "Get", url, False
'    http.send
'
'    url = Replace(http.responseText, ",", ":")
'    Dim responseJSON As String
'    responseJSON = GetDataFromURL_APP("192.168.29.182:5400", "POST", "Select * from Ord_mast where status_flag = '0'")
'    If Len(responseJSON) < 25 Then
'        MsgBox responseJSON, , "EzBiz"
'        Exit Sub
'    End If
'
'    Dim p
'    Set p = JSON.parse(responseJSON)
'    Screen.MousePointer = vbNormal
'    MsgBox p.Item("message"), , "EzBiz"
'    Screen.MousePointer = vbHourglass
'
'    Dim r As Long
'    If p.Item("status") = 1 Then
'        MsgBox ""
'    Else
'        Exit Sub
'    End If
'    Screen.MousePointer = vbNormal

    
    
    Exit Sub
''    Dim authKey, URL, mobiles, senderId, smsContentType, Message, groupId, signature, routeId As String
''
''    Dim V_signature, V_groupId, scheduleddate, V_scheduleddate As String
''
''    Dim objXML As Object
''
''    Dim getDataString As String
''
''
''
''    authKey = "79b6b4fa19ce856a0e0833bcad9d28b" '"Sample Auth key" 'eg -- 16 digits alphanumeric
''
''    URL = "http://loginsms.ewyde.com/rest/services/sendSMS/sendGroupSms"
''
''    mobiles = "9072999927" '"Sample" '99999999xx,99999998xx
''
''    smsContentType = "english" 'eg - english or unicode
''
''    Message = "Hello this is test"
''
''    senderId = "EWYSMS" '"Sample" 'eg -- Testin'
''
''    routeId = 11 '"Sample" 'eg 1'
''
'''    scheduleddate = "" 'optional if(scheduledate  eg “26/08/2015 17:00”)
'''
'''    signature = "" 'optional if(signature available  eg “1”)
'''
'''    groupId = "" 'optional if(groupId available eg “1”)
''
''
''
''    If (Len(signature) > 0) Then
''
''        V_signature = "&signature=" & signature
''
''    End If
''
''    If (Len(groupId) > 0) Then
''
''        V_groupId = "&groupId=" & groupId
''
''    End If
''
''    If (Len(scheduleddate) > 0) Then
''
''        V_scheduleddate = "&scheduleddate=" & scheduleddate
''
''    End If
''
''    'URL = "http://loginsms.ewyde.com/rest/services/sendSMS/sendGroupSms" '"Sample"
''
''    'use for API Reference
''
''    Set objXML = CreateObject("Microsoft.XMLHTTP")
''
''    'creating data for url
''
''    getDataString = URL + "AUTH_KEY=" + authKey + "&message=" + Message + "&mobileNos=" + mobiles + "&senderId=" + senderId + "&routeId=" + routeId + "&smsContentType=" + smsContentType + V_scheduleddate + V_groupId + V_signature
''
''    objXML.Open "GET", getDataString, False
''
''    objXML.send
''
''     If Len(objXML.responseText) > 0 Then
''
''            MsgBox objXML.responseText
''
''     End If

End Sub
 

Function URLEncode(ByVal Text As String) As String

    Dim i As Integer

    Dim acode As Integer

    Dim Char As String

    

    URLEncode = Text

    

    For i = Len(URLEncode) To 1 Step -1

        acode = Asc(Mid$(URLEncode, i, 1))

        Select Case acode

            Case 48 To 57, 65 To 90, 97 To 122

                ' don't touch alphanumeric chars

            Case 32

                ' replace space with "+"

                Mid$(URLEncode, i, 1) = "+"

            Case Else

                ' replace punctuation chars with "%hex"

                URLEncode = Left$(URLEncode, i - 1) & "%" & Hex$(acode) & Mid$(URLEncode, i + 1)

        End Select

    Next

    

End Function

Private Sub CmdSmry_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    FRMSTKSUMMRY.Show
    FRMSTKSUMMRY.SetFocus
End Sub

Private Sub CmdStaff_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmDStkMvmnt_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    FrmStkmovmnt.Show
    FrmStkmovmnt.SetFocus
End Sub

Private Sub CmdSearch_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    FrmPriceAnalysis.Show
    FrmPriceAnalysis.SetFocus
End Sub

Private Sub CmdQutn_Click()
    'Exit Sub
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From QTNMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(MDIMAIN.lblec.Caption))
        Exit Sub
    End If
    FRMQUOTATION.Show
    FRMQUOTATION.SetFocus
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description, , "EzBiz"
End Sub

Private Sub CMDPRETURN_Click()
    FRMPURCHASERET.Show
    FRMPURCHASERET.SetFocus
End Sub

Private Sub CMDPRETURN_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdProduction_Click()
    FRMRAWMIX.Show
    FRMRAWMIX.SetFocus
End Sub

Private Sub CmdRetailBill_Click()

    '    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS1.Show
            FRMPOS1.SetFocus
        Else
            frmsales.Show
            frmsales.SetFocus
        End If
    Else
        If frmLogin.rs!Level = "5" Then
            FRMPOS1.Show
            FRMPOS1.SetFocus
        Else
            If SALESLT_FLAG = "Y" Then
                FRMGSTRSM1.Show
                FRMGSTRSM1.SetFocus
            Else
                FRMGSTR.Show
                FRMGSTR.SetFocus
            End If
        End If
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
    
        
End Sub

Private Sub CmdRetailBill2_Click()
   
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS2.Show
            FRMPOS2.SetFocus
        Else
            FRMSALES1.Show
            FRMSALES1.SetFocus
        End If
    Else
        If frmLogin.rs!Level = "5" Then
            FRMPOS2.Show
            FRMPOS2.SetFocus
        Else
            If SALESLT_FLAG = "Y" Then
                FRMGSTRSM2.Show
                FRMGSTRSM2.SetFocus
            Else
                FRMGSTR1.Show
                FRMGSTR1.SetFocus
            End If
        End If
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
End Sub

Private Sub CmdRetailBill_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdRetailBill2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CMDSALERETURN_Click()
    frmSalesReturn.Show
    frmSalesReturn.SetFocus
End Sub

Private Sub CMDSALERETURN_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmDStkMvmnt_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdSupMast_Click()
    frmsuppliermast.Show
    frmsuppliermast.SetFocus
End Sub

Private Sub CmdSupMast_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdTransfer_Click()
    FRMTRANSFER.Show
    FRMTRANSFER.SetFocus
End Sub

Private Sub CMDTRANSFER_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CMDTRANSFER_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CmdVsale_Click()
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
        FRMSALESM.Show
        FRMSALESM.SetFocus
    Else
        FRMGSTBR.Show
        FRMGSTBR.SetFocus
    End If
End Sub

Private Sub Command1_Click()
    Dim RemDrive_Avail As Boolean
    Dim r&, allDrives$, D, stickid
    Dim aronedrive() As String
    
    RemDrive_Avail = False
    allDrives$ = Space$(64)
    r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
    allDrives$ = Left$(allDrives$, r&)
    aronedrive = Split(allDrives, vbNullChar)
    For D = UBound(aronedrive) To 1 Step -1
        If GetDriveType(aronedrive(D)) = DRIVE_REMOVABLE Then
            stickid = aronedrive(D):
            RemDrive_Avail = True
            Exit For
        End If
        'If getdrivetype(aronedrive(d)) = DRIVE_REMOTE Then stickid = aronedrive(d): Exit For
    Next
    If RemDrive_Avail = False Then
        MsgBox "No Flash drive Connected", vbOKOnly, "Backup"
        Exit Sub
    End If
    Call backup_database(aronedrive(D))
    
    Exit Sub
'    Dim RSTITEMMAST, rststock As ADODB.Recordset
'    Dim INWARD, OUTWARD As Double
'
    'db.Execute "delete FROM RTRXFILE WHERE TRX_TYPE ='OP'"
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM CUSTMAST WHERE KGST = '123'", db, adOpenStatic, adLockOptimistic, adCmdText
'    If Not (RSTITEMMAST.EOF And RSTITEMMAST.BOF) Then
'        If IsNull(RSTITEMMAST!KGST) Or RSTITEMMAST!KGST = "" Then
'            RSTITEMMAST!KGST = ""
'            RSTITEMMAST.Update
'        End If
'    End If
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
        
'    On Error GoTo eRRHAND
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM ITEMMAST ", db, adOpenStatic, adLockOptimistic
'    Do Until rststock.EOF
'        rststock!OPEN_QTY = 0
'        rststock.Update
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
'
'
'    Screen.MousePointer = vbHourglass
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM RTRXFILE ORDER BY ITEM_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTITEMMAST.EOF
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * FROM ITEMMAST WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'        If Not (rststock.EOF And rststock.BOF) Then
'            RSTITEMMAST!Category = rststock!Category
'            RSTITEMMAST.Update
'        End If
'        rststock.Close
'        Set rststock = Nothing
'        RSTITEMMAST.MoveNext
'    Loop
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
'    Screen.MousePointer = vbNormal
'    Exit Sub
'
'    On Error GoTo eRRHAND
'    Screen.MousePointer = vbHourglass
'    Set RSTITEMMAST = New ADODB.Recordset
'    RSTITEMMAST.Open "SELECT * FROM ITEMMAST ORDER BY ITEM_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTITEMMAST.EOF
'        INWARD = 0
'        OUTWARD = 0
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'        Do Until rststock.EOF
'            INWARD = INWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
'
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockReadOnly
'        Do Until rststock.EOF
'            OUTWARD = OUTWARD + (IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK))
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
'
'        RSTITEMMAST!CLOSE_QTY = INWARD - OUTWARD
'        RSTITEMMAST!RCPT_QTY = INWARD
'        RSTITEMMAST!ISSUE_QTY = OUTWARD
'        RSTITEMMAST.Update
'        RSTITEMMAST.MoveNext
'    Loop
'    RSTITEMMAST.Close
'    Set RSTITEMMAST = Nothing
    Screen.MousePointer = vbNormal
'    FRMRETAILWO.Show
'    FRMRETAILWO.SetFocus
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
End Sub

Private Sub CommandButton1_Click()
    Dim RSTITEMMAST As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTBALQTY As ADODB.Recordset
    Dim INWARD As Double
    Dim OUTWARD As Double
    Dim BALQTY As Double
    Dim DIFFQTY As Double
    Dim i As Long
''''    db.Execute "delete from cashatrxfile"
''''    db.Execute "delete from dbtpymt"
''''    db.Execute "delete from BANK_TRX"
''''    db.Execute "delete from CATEGORY"
''''    Exit Sub
    
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    On Error GoTo ERRHAND
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    db.Execute "Update RTRXFILE set BAL_QTY = 0 WHERE ISNULL(BAL_QTY) OR BAL_QTY <0 "
    Set RSTITEMMAST = New ADODB.Recordset
    RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(UN_BILL) OR UN_BILL <> 'Y') AND ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME DESC", db, adOpenStatic, adLockOptimistic, adCmdText
    RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
    MDIMAIN.vbalProgressBar1.Max = RSTITEMMAST.RecordCount
    Do Until RSTITEMMAST.EOF
        
        BALQTY = 0
        Set RSTBALQTY = New ADODB.Recordset
        RSTBALQTY.Open "Select SUM(BAL_QTY) FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "'", db, adOpenForwardOnly
        If Not (RSTBALQTY.EOF And RSTBALQTY.BOF) Then
            BALQTY = IIf(IsNull(RSTBALQTY.Fields(0)), 0, RSTBALQTY.Fields(0))
        End If
        RSTBALQTY.Close
        Set RSTBALQTY = Nothing
        
        INWARD = 0
        OUTWARD = 0
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(QTY + FREE_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
            
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT SUM(FREE_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
'        If Not (rststock.EOF And rststock.BOF) Then
'            INWARD = INWARD + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
'        End If
'        rststock.Close
'        Set rststock = Nothing
        
                
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until rststock.EOF
'            OUTWARD = OUTWARD + IIf(IsNull(rststock!QTY), 0, rststock!QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
'            OUTWARD = OUTWARD + IIf(IsNull(rststock!FREE_QTY), 0, rststock!FREE_QTY) * IIf(IsNull(rststock!LOOSE_PACK), 1, rststock!LOOSE_PACK)
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "Select SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly, adCmdText
        If Not (rststock.EOF And rststock.BOF) Then
            OUTWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
'        Set rststock = New ADODB.Recordset
'        rststock.Open "Select SUM(FREE_QTY * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI' OR TRX_TYPE='GI' OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR') ", db, adOpenStatic, adLockReadOnly, adCmdText
'        If Not (rststock.EOF And rststock.BOF) Then
'            OUTWARD = OUTWARD + IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
'        End If
'        rststock.Close
'        Set rststock = Nothing
'
        If Round(INWARD - OUTWARD, 2) = Round(BALQTY, 2) Then GoTo SKIP_BALCHECK
        
        
        db.Execute "Update RTRXFILE set BAL_QTY = QTY where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' "
        BALQTY = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * FROM RTRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rststock.EOF
            BALQTY = 0
            Set RSTBALQTY = New ADODB.Recordset
            RSTBALQTY.Open "Select SUM(QTY) FROM TRXSUB WHERE R_TRX_YEAR ='" & rststock!TRX_YEAR & "' AND R_TRX_TYPE='" & rststock!TRX_TYPE & "' AND R_VCH_NO = " & rststock!VCH_NO & " AND R_LINE_NO = " & rststock!LINE_NO & "", db, adOpenForwardOnly
            If Not (RSTBALQTY.EOF And RSTBALQTY.BOF) Then
                BALQTY = IIf(IsNull(RSTBALQTY.Fields(0)), 0, RSTBALQTY.Fields(0))
            End If
            RSTBALQTY.Close
            Set RSTBALQTY = Nothing
            
            rststock!BAL_QTY = rststock!BAL_QTY - BALQTY
            rststock.Update
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        
        
        db.Execute "Update RTRXFILE set BAL_QTY = 0 where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY <0"
        BALQTY = 0
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT SUM(BAL_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
        If Not (rststock.EOF And rststock.BOF) Then
            BALQTY = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
        End If
        rststock.Close
        Set rststock = Nothing
        
        If Round(INWARD - OUTWARD, 2) < BALQTY Then
            DIFFQTY = BALQTY - (Round(INWARD - OUTWARD, 2))
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND BAL_QTY > 0 ORDER BY VCH_DATE ", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                If DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) >= 0 Then
                    DIFFQTY = DIFFQTY - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY)
                    rststock!BAL_QTY = 0
                    rststock.Update
                Else
                    rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) - DIFFQTY, 2)
                    DIFFQTY = 0
                    rststock.Update
                End If
                If DIFFQTY <= 0 Then Exit Do
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
        ElseIf Round(INWARD - OUTWARD, 2) > BALQTY Then
            DIFFQTY = Round((INWARD - OUTWARD), 2) - BALQTY
            Set rststock = New ADODB.Recordset
            rststock.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockOptimistic, adCmdText
            Do Until rststock.EOF
                If DIFFQTY <= IIf(IsNull(rststock!QTY), 0, rststock!QTY) - IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) Then
                    rststock!BAL_QTY = Round(IIf(IsNull(rststock!BAL_QTY), 0, rststock!BAL_QTY) + DIFFQTY, 2)
                    DIFFQTY = 0
                Else
                    If Not rststock!BAL_QTY = IIf(IsNull(rststock!QTY), 0, rststock!QTY) Then
                        rststock!BAL_QTY = Round(IIf(IsNull(rststock!QTY), 0, rststock!QTY), 2)
                        DIFFQTY = DIFFQTY - IIf(IsNull(rststock!QTY), 0, rststock!QTY)
                    End If
                End If
                rststock.Update
                If DIFFQTY <= 0 Then Exit Do
                rststock.MoveNext
            Loop
            rststock.Close
            Set rststock = Nothing
            'MsgBox ""
        End If
        
SKIP_BALCHECK:
        RSTITEMMAST!CLOSE_QTY = Round(INWARD - OUTWARD, 2)
        RSTITEMMAST!RCPT_QTY = INWARD
        RSTITEMMAST!ISSUE_QTY = OUTWARD
        RSTITEMMAST.Update
        RSTITEMMAST.MoveNext
        MDIMAIN.vbalProgressBar1.Value = MDIMAIN.vbalProgressBar1.Value + 1
    Loop
    RSTITEMMAST.Close
    Set RSTITEMMAST = Nothing
    
    Screen.MousePointer = vbNormal
    MDIMAIN.vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    MDIMAIN.vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub CommandButton1_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CommandButton2_Click()
    FRMContra.Show
    FRMContra.SetFocus
End Sub

Private Sub CommandButton2_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CommandButton1_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub CommandButton3_Click()
    FRMJournal.Show
    FRMJournal.SetFocus
End Sub

Private Sub CommandButton4_Click()
    FRMRAWMIX2.Show
    FRMRAWMIX2.SetFocus
End Sub

Private Sub CommandButton4_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub CommandButton4_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub Label1_Click()

End Sub

Private Sub MDIForm_Activate()
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        mnuwsale1.Visible = False
        mnuwsale.Visible = False
        mnub2c1.Caption = "Sales -1"
        mnu2.Caption = "Sales -2"
        mnuretail.Caption = "Sales -3"
        CMDBILLRETAIL.Caption = "Service Bill"
        CmdVsale.Caption = "Sales Bill (Counter)"
    End If
    If LBLAMC.Caption = "Y" Then
        FrmAMC.Show
        FrmAMC.SetFocus
        If FrmAMC.GRDSTOCK.rows <= 1 Then Unload FrmAMC
    End If
End Sub

Private Sub MDIForm_Load()
 'CLOSEALL = 1
    If frmLogin.rs!Level = "1" Then 'Or frmLogin.rs!Level = "4" Then   'MANAGER & SECOND ADMIN PRIVILAGE
        MNUUSER.Visible = False
        STKMVMNT.Visible = False
        MNUCOMSN.Visible = False
        mnusalesrep.Visible = False
'        MNUSHOPINFO.Visible = False
'        MNUACCOUNTS.Visible = False
'        cmdpurchase.Visible = False
'        CMDPRETURN.Visible = False
'        MNUSUPPLIER.Visible = False
'        MNUPO.Visible = False
'        mnuOP.Visible = False
'        MnuExpenseSt.Visible = False
'        MNUEXPENTRY.Visible = False
'        MnuExpStaff.Visible = False
'        PGAYR.Visible = False
'        PR.Visible = False
'        mnupymntdue.Visible = False
'        MnuExpReg.Visible = False
'        MnuExpStaffReg.Visible = False
'        MNUP_RETURNREG.Visible = False
    ElseIf frmLogin.rs!Level = "4" Then 'SECOND ADMIN PRIVILAGE
        MNUUSER.Visible = False
    ElseIf frmLogin.rs!Level = "2" Or frmLogin.rs!Level = "5" Then  'SALES
        'MDIMAIN.mnuloc.Visible = False
        'cmdorder.Visible = False
        MNUUSER.Visible = False
        STKMVMNT.Visible = False
        cmdpurchase.Visible = False
        CMDPRETURN.Visible = False
        CmdProduction.Visible = False
        CommandButton4.Visible = False
        CMD62.Visible = False
        CmdTransfer.Visible = False
        CmdReceipt.Visible = False
        CmdPayment.Visible = False
        'CmdExp.Visible = False
        CmdStaff.Visible = False
        CmdLend.Visible = False
        CmDIncome.Visible = False
        CmdJournal.Visible = False
        CommandButton2.Visible = False 'CONTRA
        CmdBook.Visible = False
        CmdSupMast.Visible = False
        MNUCOMSN.Visible = False
        MNUSHOPINFO.Visible = False
        mnuaccounts.Visible = False
        mnupurch.Visible = False
        MNUEMPLOYEE.Visible = False
        MNULEND.Visible = False
        MNUSUPPLIER.Visible = False
        MNUPO.Visible = False
        MNUOP.Visible = False
        MnuExpenseSt.Visible = False
        MNUEXPENTRY.Visible = False
        PGAYR.Visible = False
'        PR.Visible = False
        mnupymntdue.Visible = False
        mnuExpReg.Visible = False
        MnuExpStaffReg.Visible = False
        MNUP_RETURNREG.Visible = False
        MNUREPORT.Visible = False
        mnuvoucher.Visible = False
        mnuhead.Visible = False
        MnuMinQty.Visible = False
        mnuOPCash.Visible = False
        MnuSer_Stk.Visible = False
        MNUPO.Visible = False
        MNUOP.Visible = False
        mnustktransfer.Visible = False
        MNUDAMAGE.Visible = False
        MNUGIFT.Visible = False
        MNUBANKBOOK.Visible = False
        MNUFMLA.Visible = False
        
        If frmLogin.rs!Level = "5" Then
            CmdQutn.Visible = False
            MNUQTN.Visible = False
            CMDSALERETURN.Visible = False
        End If
        CmdItemMast.Visible = False
        CmdSmry.Visible = False
        CMDDELIVERY.Visible = False
        CmdVsale.Visible = False
        MNUYEAR.Visible = False
        mnumerge.Visible = False
        mnumergemulti.Visible = False
        mnutblewise.Visible = False
        mnustkcr.Visible = False
        mnustk.Visible = False
        MNUOPSTK.Visible = False
        mnuproduction.Visible = False
        mnuimport.Visible = False
        MNUDAMAGE.Visible = False
        mnudmgbr.Visible = False
        MNUPROCESS.Visible = False
        MNUYEAR.Visible = False
    ElseIf frmLogin.rs!Level = "3" Then
        'MDIMAIN.mnuloc.Visible = False
        MNUUSER.Visible = False
        STKMVMNT.Visible = False
        CmdRetailBill.Visible = False
        CmdRetailBill2.Visible = False
        CmdGST.Visible = False
        CMDSALERETURN.Visible = False
        CMDBILLRETAIL.Visible = False
        CMDDELIVERY.Visible = False
        CmdTransfer.Visible = False
        
        CmdPayment.Visible = False
        CmdReceipt.Visible = False
        CmdExp.Visible = False
        CmdStaff.Visible = False
        CmdLend.Visible = False
        CmDIncome.Visible = False
        CmdJournal.Visible = False
        CommandButton2.Visible = False 'CONTRA
        CmdBook.Visible = False
        CmdCust.Visible = False
        CmdQutn.Visible = False
        cmdorder.Visible = False
        CmdExpiry.Visible = False
        
        MNUCOMSN.Visible = False
        MNUSHOPINFO.Visible = False
        mnusale.Visible = False
        MNUEMPLOYEE.Visible = False
        MNULEND.Visible = False
        MNUCUST.Visible = False
        MNUOP.Visible = False
        MnuExpenseSt.Visible = False
        MNUEXPENTRY.Visible = False
        PGAYR.Visible = False
        MNUQTN.Visible = False
        mnotkorder.Visible = False
'        PR.Visible = False
        mnupymntdue.Visible = False
        mnuExpReg.Visible = False
        MnuExpStaffReg.Visible = False
        MNUP_RETURNREG.Visible = False
        
        
        MNUREPORT.Visible = False
        mnuvoucher.Visible = False
        MNUREPORT.Visible = False
        mnuvoucher.Visible = False
        mnuhead.Visible = False
        MnuMinQty.Visible = False
        mnuOPCash.Visible = False
        MnuSer_Stk.Visible = False
        MNUOP.Visible = False
        mnustktransfer.Visible = False
        MNUDAMAGE.Visible = False
        MNUGIFT.Visible = False
        MNUBANKBOOK.Visible = False
        MNUFMLA.Visible = False
        CmDStkMvmnt.Visible = False

'        MNUSHOPINFO.Visible = False
'        MNUACCOUNTS.Visible = False
'        CMDBILLRETAIL.Visible = False
'        CMDSALERETURN.Visible = False
'        CommandButton2.Visible = False
'        MNUCUST.Visible = False
'        MnuExpenseSt.Visible = False
'        MNUEXPENTRY.Visible = False
'        MnuExpStaff.Visible = False
'        MNURCPTREG.Visible = False
'        SR.Visible = False
'        mnurcptdue.Visible = False
'        MnuExpReg.Visible = False
'        MnuExpStaffReg.Visible = False
'        MNUDEL_REG.Visible = False
'        MNUS_RETURNREG.Visible = False
'        FRMROUTE.Visible = False
    End If
'    FrmOccassion.Show
'    FrmOccassion.SetFocus
'    If FrmOccassion.GRDSTOCK.Rows <= 1 Then Unload FrmOccassion
    
'    If Month(Date) >= 5 And Year(Date) >= 2021 And Day(Date) > 13 Then
'        db.Execute "delete From USERS "
'    End If
    
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
'        Call backup_database("D:\Backup\")
    Select Case MsgBox("DO YOU WANT TO TAKE BACKUP... THIS MAY TAKE SEVERAL MINUTES TO FINISH!!! MAKE SURE THE FLASH DRIVE IS CONNECTED" & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNoCancel)
        Case vbYes
            Dim RemDrive_Avail As Boolean
            Dim r&, allDrives$, D, stickid
            Dim aronedrive() As String

            RemDrive_Avail = False
            allDrives$ = Space$(64)
            r& = GetLogicalDriveStrings(Len(allDrives$), allDrives$)
            allDrives$ = Left$(allDrives$, r&)
            aronedrive = Split(allDrives, vbNullChar)
            For D = UBound(aronedrive) To 1 Step -1
                If GetDriveType(aronedrive(D)) = DRIVE_REMOVABLE Then
                    stickid = aronedrive(D):
                    RemDrive_Avail = True
                    Exit For
                End If
                'If getdrivetype(aronedrive(d)) = DRIVE_REMOTE Then stickid = aronedrive(d): Exit For
            Next
            If RemDrive_Avail = False Then
                If MsgBox("No Flash drive Connected" & Chr(13) & "Do you want to take backup on external device", vbYesNo, "Backup") = vbYes Then
                    Cancel = 1
                    fRMbackup.Show
                    fRMbackup.SetFocus
                    Exit Sub
                End If
                Cancel = 1
                Exit Sub
            End If
        
'            Dim RSTITEMMAST As ADODB.Recordset
'            Dim rststock As ADODB.Recordset
'            Dim RSTBALQTY As ADODB.Recordset
'            Dim INWARD As Double
'            Dim OUTWARD As Double
'
'
'            Screen.MousePointer = vbHourglass
'            On Error GoTo ERRHAND
'            Set RSTITEMMAST = New ADODB.Recordset
'            RSTITEMMAST.Open "SELECT * FROM ITEMMAST WHERE ucase(CATEGORY) <> 'SELF' AND ucase(CATEGORY) <> 'SERVICE CHARGE' AND ucase(CATEGORY) <> 'SERVICES' ORDER BY ITEM_NAME", db, adOpenStatic, adLockOptimistic, adCmdText
'            RSTITEMMAST.Properties("Update Criteria").Value = adCriteriaKey
'            Do Until RSTITEMMAST.EOF
'
'                INWARD = 0
'                OUTWARD = 0
'                Set rststock = New ADODB.Recordset
'                rststock.Open "SELECT SUM(QTY + FREE_QTY) FROM RTRXFILE where ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' ", db, adOpenForwardOnly
'                If Not (rststock.EOF And rststock.BOF) Then
'                    INWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'                Set rststock = New ADODB.Recordset
'                rststock.Open "Select SUM((QTY + FREE_QTY) * LOOSE_PACK) FROM TRXFILE WHERE  ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "' AND ((TRX_TYPE='HI' AND CST =0) OR (TRX_TYPE='GI' AND CST =0) OR (TRX_TYPE='SV' AND CST =0) OR (TRX_TYPE='SI' AND CST =0) OR (TRX_TYPE='RI' AND CST =0) OR (TRX_TYPE='WO' AND CST =0) OR (TRX_TYPE='VI' AND CST =0) OR (TRX_TYPE='TF' AND CST =0) OR TRX_TYPE='DN' OR TRX_TYPE='WP' OR TRX_TYPE='PR' OR TRX_TYPE='RM' OR TRX_TYPE='PC' OR TRX_TYPE='MI'  OR TRX_TYPE='DG' OR TRX_TYPE='DM' OR TRX_TYPE='GF' OR TRX_TYPE='RW' OR TRX_TYPE='SR' OR TRX_TYPE='EP' OR TRX_TYPE='EX' OR TRX_TYPE='RM' OR TRX_TYPE='PC') ", db, adOpenStatic, adLockReadOnly, adCmdText
'                If Not (rststock.EOF And rststock.BOF) Then
'                    OUTWARD = IIf(IsNull(rststock.Fields(0)), 0, rststock.Fields(0))
'                End If
'                rststock.Close
'                Set rststock = Nothing
'
'
'                RSTITEMMAST!CLOSE_QTY = Round(INWARD - OUTWARD, 2)
'                RSTITEMMAST!RCPT_QTY = INWARD
'                RSTITEMMAST!ISSUE_QTY = OUTWARD
'                RSTITEMMAST.Update
'                RSTITEMMAST.MoveNext
'            Loop
'            RSTITEMMAST.Close
'            Set RSTITEMMAST = Nothing
            
            Call backup_database(aronedrive(D))
        Case vbCancel
            Cancel = 1
            Exit Sub
        Case vbNo
            On Error GoTo ERRHAND
            Dim Strconnct As String
            Dim db2 As New ADODB.Connection
            
            Screen.MousePointer = vbHourglass
            If Dir(App.Path & "\Backup", vbDirectory) = "" Then MkDir App.Path & "\Backup"
            If Not FileExists(App.Path & "\mysqldump.exe") Then
                Screen.MousePointer = vbNormal
                MsgBox "File not exists", , "EzBiz"
                Exit Sub
            End If
            Dim cmd As String
            Dim strBackupEXT As String
            strBackupEXT = "bk" & Format(Format(Date, "ddmmyy"), "000000") & Format(Format(Time, "HHMMSS"), "")
            DoEvents
            'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " > " & App.Path & "\Backup\" & strBackupEXT
            cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " > " & App.Path & "\Backup\" & strBackupEXT
            Call execCommand(cmd)
            
            If DBPath = "localhost" Then
                db.Execute "DROP DATABASE if exists tempdb "
                db.Execute "CREATE DATABASE tempdb;"
                
                Strconnct = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=tempdb;User=root; Password=###%%database%%###ret; Option=2;"
                db2.Open Strconnct
                db2.CursorLocation = adUseClient
        
                
                
                'db2.Execute "SHOW DATABASES;"
                
        
                'Strconnct = "Driver={MySQL ODBC 5.1 Driver};Server=localhost;Port=3306;Database=tempdb;User=root; Password=###%%database%%###ret; Option=2;"
                Set db2 = New ADODB.Connection
                db2.Open Strconnct
                db2.CursorLocation = adUseClient
                        
                db2.Execute "CREATE TABLE `actmast` LIKE " & dbase1 & ".`actmast`"
                db2.Execute "INSERT INTO `actmast` SELECT * FROM " & dbase1 & ".`actmast`"
'
'                db2.Execute "CREATE TABLE `act_ky` LIKE " & dbase1 & ".`act_ky`"
'                db2.Execute "INSERT INTO `act_ky` SELECT * FROM " & dbase1 & ".`act_ky`"
'
'                db2.Execute "CREATE TABLE `address_book` LIKE " & dbase1 & ".`address_book`"
'                db2.Execute "INSERT INTO `address_book` SELECT * FROM " & dbase1 & ".`address_book`"
'
'                db2.Execute "CREATE TABLE `arealist` LIKE " & dbase1 & ".`arealist`"
'                db2.Execute "INSERT INTO `arealist` SELECT * FROM " & dbase1 & ".`arealist`"
'
'                db2.Execute "CREATE TABLE `atrxfile` LIKE " & dbase1 & ".`atrxfile`"
'                db2.Execute "INSERT INTO `atrxfile` SELECT * FROM " & dbase1 & ".`atrxfile`"
'
'                db2.Execute "CREATE TABLE `atrxsub` LIKE " & dbase1 & ".`atrxsub`"
'                db2.Execute "INSERT INTO `atrxsub` SELECT * FROM " & dbase1 & ".`atrxsub`"
'
'                db2.Execute "CREATE TABLE `bankcode` LIKE " & dbase1 & ".`bankcode`"
'                db2.Execute "INSERT INTO `bankcode` SELECT * FROM " & dbase1 & ".`bankcode`"
'
'                db2.Execute "CREATE TABLE `bankletters` LIKE " & dbase1 & ".`bankletters`"
'                db2.Execute "INSERT INTO `bankletters` SELECT * FROM " & dbase1 & ".`bankletters`"
'
'                db2.Execute "CREATE TABLE `bank_trx` LIKE " & dbase1 & ".`bank_trx`"
'                db2.Execute "INSERT INTO `bank_trx` SELECT * FROM " & dbase1 & ".`bank_trx`"
'
'                db2.Execute "CREATE TABLE `billdetails` LIKE " & dbase1 & ".`billdetails`"
'                db2.Execute "CREATE TABLE `bonusmast` LIKE " & dbase1 & ".`bonusmast`"
'                'db2.Execute "CREATE TABLE `bookfile` LIKE " & dbase1 & ".`bookfile`"
'                'db2.Execute "CREATE TABLE `cancinv` LIKE " & dbase1 & ".`cancinv`"
'
'                db2.Execute "CREATE TABLE `cashatrxfile` LIKE " & dbase1 & ".`cashatrxfile`"
'                db2.Execute "INSERT INTO `cashatrxfile` SELECT * FROM " & dbase1 & ".`cashatrxfile`"
'
'                db2.Execute "CREATE TABLE `category` LIKE " & dbase1 & ".`category`"
'                db2.Execute "INSERT INTO `category` SELECT * FROM " & dbase1 & ".`category`"
'
'                db2.Execute "CREATE TABLE `chqmast` LIKE " & dbase1 & ".`chqmast`"
'
'                db2.Execute "CREATE TABLE `compinfo` LIKE " & dbase1 & ".`compinfo`"
'                db2.Execute "INSERT INTO `compinfo` SELECT * FROM " & dbase1 & ".`compinfo`"
'
'                db2.Execute "CREATE TABLE `cont_mast` LIKE " & dbase1 & ".`cont_mast`"
'
                db2.Execute "CREATE TABLE `crdtpymt` LIKE " & dbase1 & ".`crdtpymt`"
                db2.Execute "INSERT INTO `crdtpymt` SELECT * FROM " & dbase1 & ".`crdtpymt`"
                
                db2.Execute "CREATE TABLE `custmast` LIKE " & dbase1 & ".`custmast`"
                db2.Execute "INSERT INTO `custmast` SELECT * FROM " & dbase1 & ".`custmast`"
                
'                db2.Execute "CREATE TABLE `custtrxfile` LIKE " & dbase1 & ".`custtrxfile`"
'                db2.Execute "INSERT INTO `custtrxfile` SELECT * FROM " & dbase1 & ".`custtrxfile`"
'
'                db2.Execute "CREATE TABLE `cust_details` LIKE " & dbase1 & ".`cust_details`"
'                db2.Execute "INSERT INTO `cust_details` SELECT * FROM " & dbase1 & ".`cust_details`"
'
'                db2.Execute "CREATE TABLE `damaged` LIKE " & dbase1 & ".`damaged`"
'                db2.Execute "INSERT INTO `damaged` SELECT * FROM " & dbase1 & ".`damaged`"
'
                db2.Execute "CREATE TABLE `dbtpymt` LIKE " & dbase1 & ".`dbtpymt`"
                db2.Execute "INSERT INTO `dbtpymt` SELECT * FROM " & dbase1 & ".`dbtpymt`"
'
'                db2.Execute "CREATE TABLE `de_ret_details` LIKE " & dbase1 & ".`de_ret_details`"
'
'                db2.Execute "CREATE TABLE `expiry` LIKE " & dbase1 & ".`expiry`"
'                db2.Execute "CREATE TABLE `explist` LIKE " & dbase1 & ".`explist`"
'                db2.Execute "CREATE TABLE `expsort` LIKE " & dbase1 & ".`expsort`"
'                db2.Execute "CREATE TABLE `fqtylist` LIKE " & dbase1 & ".`fqtylist`"
'
'                db2.Execute "CREATE TABLE `gift` LIKE " & dbase1 & ".`gift`"
'                db2.Execute "INSERT INTO `gift` SELECT * FROM " & dbase1 & ".`gift`"
'
'                db2.Execute "CREATE TABLE `hsn_trxfile` LIKE " & dbase1 & ".`hsn_trxfile`"
'                db2.Execute "CREATE TABLE `hsn_trxmast` LIKE " & dbase1 & ".`hsn_trxmast`"
'
'
'                db2.Execute "CREATE TABLE `manufact` LIKE " & dbase1 & ".`manufact`"
'                db2.Execute "INSERT INTO `manufact` SELECT * FROM " & dbase1 & ".`manufact`"
'
'
'                db2.Execute "CREATE TABLE `ordissue` LIKE " & dbase1 & ".`ordissue`"
'                db2.Execute "CREATE TABLE `ordsub` LIKE " & dbase1 & ".`ordsub`"
'                db2.Execute "CREATE TABLE `password` LIKE " & dbase1 & ".`password`"
'                db2.Execute "CREATE TABLE `passwords` LIKE " & dbase1 & ".`passwords`"
'
'                db2.Execute "CREATE TABLE `pomast` LIKE " & dbase1 & ".`pomast`"
'                db2.Execute "INSERT INTO `pomast` SELECT * FROM " & dbase1 & ".`pomast`"
'
'                db2.Execute "CREATE TABLE `posub` LIKE " & dbase1 & ".`posub`"
'                db2.Execute "INSERT INTO `posub` SELECT * FROM " & dbase1 & ".`posub`"
'
'                db2.Execute "CREATE TABLE `pricetable` LIKE " & dbase1 & ".`pricetable`"
'
'                db2.Execute "CREATE TABLE `purcahsereturn` LIKE " & dbase1 & ".`purcahsereturn`"
'                db2.Execute "INSERT INTO `purcahsereturn` SELECT * FROM " & dbase1 & ".`purcahsereturn`"
'
'                db2.Execute "CREATE TABLE `purch_return` LIKE " & dbase1 & ".`purch_return`"
'                db2.Execute "INSERT INTO `purch_return` SELECT * FROM " & dbase1 & ".`purch_return`"
'
'                db2.Execute "CREATE TABLE `qtnmast` LIKE " & dbase1 & ".`qtnmast`"
'                db2.Execute "INSERT INTO `qtnmast` SELECT * FROM " & dbase1 & ".`qtnmast`"
'
'                db2.Execute "CREATE TABLE `qtnsub` LIKE " & dbase1 & ".`qtnsub`"
'                db2.Execute "INSERT INTO `qtnsub` SELECT * FROM " & dbase1 & ".`qtnsub`"
'
'                db2.Execute "CREATE TABLE `reorder` LIKE " & dbase1 & ".`reorder`"
'                db2.Execute "CREATE TABLE `replcn` LIKE " & dbase1 & ".`replcn`"
'                db2.Execute "CREATE TABLE `returnmast` LIKE " & dbase1 & ".`returnmast`"
'
'                db2.Execute "CREATE TABLE `salereturn` LIKE " & dbase1 & ".`salereturn`"
'                db2.Execute "CREATE TABLE `salesledger` LIKE " & dbase1 & ".`salesledger`"
'                db2.Execute "CREATE TABLE `salesman` LIKE " & dbase1 & ".`salesman`"
'                db2.Execute "CREATE TABLE `salesreg` LIKE " & dbase1 & ".`salesreg`"
'                db2.Execute "CREATE TABLE `seldist` LIKE " & dbase1 & ".`seldist`"
'                db2.Execute "CREATE TABLE `service_stk` LIKE " & dbase1 & ".`service_stk`"
'                db2.Execute "CREATE TABLE `slip_reg` LIKE " & dbase1 & ".`slip_reg`"
'                db2.Execute "CREATE TABLE `srtrxfile` LIKE " & dbase1 & ".`srtrxfile`"
'                db2.Execute "CREATE TABLE `stockreport` LIKE " & dbase1 & ".`stockreport`"
'
'                db2.Execute "CREATE TABLE `tempcn` LIKE " & dbase1 & ".`tempcn`"
'                db2.Execute "CREATE TABLE `tempstk` LIKE " & dbase1 & ".`tempstk`"
'                db2.Execute "CREATE TABLE `temptrxfile` LIKE " & dbase1 & ".`temptrxfile`"
'                db2.Execute "CREATE TABLE `tmporderlist` LIKE " & dbase1 & ".`tmporderlist`"
                
                db2.Execute "CREATE TABLE `transmast` LIKE " & dbase1 & ".`transmast`"
                db2.Execute "INSERT INTO `transmast` SELECT * FROM " & dbase1 & ".`transmast`"
                
'                db2.Execute "CREATE TABLE `transsub` LIKE " & dbase1 & ".`transsub`"
'                db2.Execute "INSERT INTO `transsub` SELECT * FROM " & dbase1 & ".`transsub`"
'
'                db2.Execute "CREATE TABLE `trnxrcpt` LIKE " & dbase1 & ".`trnxrcpt`"
'                db2.Execute "INSERT INTO `trnxrcpt` SELECT * FROM " & dbase1 & ".`trnxrcpt`"
'
'                db2.Execute "CREATE TABLE `trxexpense` LIKE " & dbase1 & ".`trxexpense`"
'                db2.Execute "INSERT INTO `trxexpense` SELECT * FROM " & dbase1 & ".`trxexpense`"
'
'                db2.Execute "CREATE TABLE `trxexpmast` LIKE " & dbase1 & ".`trxexpmast`"
'                db2.Execute "INSERT INTO `trxexpmast` SELECT * FROM " & dbase1 & ".`trxexpmast`"
'
'                db2.Execute "CREATE TABLE `trxexp_mast` LIKE " & dbase1 & ".`trxexp_mast`"
'                db2.Execute "INSERT INTO `trxexp_mast` SELECT * FROM " & dbase1 & ".`trxexp_mast`"
                
                db2.Execute "CREATE TABLE `trxfile` LIKE " & dbase1 & ".`trxfile`"
                db2.Execute "INSERT INTO `trxfile` SELECT * FROM " & dbase1 & ".`trxfile`"
                
'                db2.Execute "CREATE TABLE `trxfileexp` LIKE " & dbase1 & ".`trxfileexp`"
'                db2.Execute "INSERT INTO `trxfileexp` SELECT * FROM " & dbase1 & ".`trxfileexp`"
                
'                db2.Execute "CREATE TABLE `trxfile_exp` LIKE " & dbase1 & ".`trxfile_exp`"
'                db2.Execute "INSERT INTO `trxfile_exp` SELECT * FROM " & dbase1 & ".`trxfile_exp`"
'
'                db2.Execute "CREATE TABLE `trxfile_sp` LIKE " & dbase1 & ".`trxfile_sp`"
'                db2.Execute "INSERT INTO `trxfile_sp` SELECT * FROM " & dbase1 & ".`trxfile_sp`"
'
'                db2.Execute "CREATE TABLE `trxincmast` LIKE " & dbase1 & ".`trxincmast`"
'                db2.Execute "INSERT INTO `trxincmast` SELECT * FROM " & dbase1 & ".`trxincmast`"
'
'                db2.Execute "CREATE TABLE `trxincome` LIKE " & dbase1 & ".`trxincome`"
'                db2.Execute "INSERT INTO `trxincome` SELECT * FROM " & dbase1 & ".`trxincome`"
'
                db2.Execute "CREATE TABLE `trxmast` LIKE " & dbase1 & ".`trxmast`"
                db2.Execute "INSERT INTO `trxmast` SELECT * FROM " & dbase1 & ".`trxmast`"
                
'                db2.Execute "CREATE TABLE `trxmast_sp` LIKE " & dbase1 & ".`trxmast_sp`"
'                db2.Execute "INSERT INTO `trxmast_sp` SELECT * FROM " & dbase1 & ".`trxmast_sp`"
'
'                db2.Execute "CREATE TABLE `trxsub` LIKE " & dbase1 & ".`trxsub`"
'                db2.Execute "INSERT INTO `trxsub` SELECT * FROM " & dbase1 & ".`trxsub`"
'
'                db2.Execute "CREATE TABLE `users` LIKE " & dbase1 & ".`users`"
'                db2.Execute "INSERT INTO `users` SELECT * FROM " & dbase1 & ".`users`"
'
'                db2.Execute "CREATE TABLE `vanstock` LIKE " & dbase1 & ".`vanstock`"
'                db2.Execute "CREATE TABLE `war_list` LIKE " & dbase1 & ".`war_list`"
'                db2.Execute "CREATE TABLE `war_trxfile` LIKE " & dbase1 & ".`war_trxfile`"
'                db2.Execute "CREATE TABLE `war_trxns` LIKE " & dbase1 & ".`war_trxns`"
'
'                db2.Execute "CREATE TABLE `astmast` LIKE " & dbase1 & ".`astmast`"
'                db2.Execute "INSERT INTO `astmast` SELECT * FROM " & dbase1 & ".`astmast`"
'
'                db2.Execute "CREATE TABLE `astrxfile` LIKE " & dbase1 & ".`astrxfile`"
'                db2.Execute "INSERT INTO `astrxfile` SELECT * FROM " & dbase1 & ".`astrxfile`"
'
'                db2.Execute "CREATE TABLE `astrxmast` LIKE " & dbase1 & ".`astrxmast`"
'                db2.Execute "INSERT INTO `astrxmast` SELECT * FROM " & dbase1 & ".`astrxmast`"
                
                db2.Execute "CREATE TABLE `RTRXFILE` LIKE " & dbase1 & ".`RTRXFILE`"
                db2.Execute "INSERT INTO `RTRXFILE` SELECT * FROM " & dbase1 & ".`RTRXFILE` "
                
                db2.Execute "CREATE TABLE `itemmast` LIKE " & dbase1 & ".`itemmast`"
                db2.Execute "INSERT INTO `itemmast` SELECT * FROM " & dbase1 & ".`itemmast` "
                
'                On Error Resume Next
'                db2.Execute "CREATE TABLE `DAMAGE_MAST` LIKE " & dbase1 & ".`DAMAGE_MAST`"
'                db2.Execute "INSERT INTO `DAMAGE_MAST` SELECT * FROM " & dbase1 & ".`DAMAGE_MAST`"
'
'                db2.Execute "CREATE TABLE `nonrcvd` LIKE " & dbase1 & ".`nonrcvd`"
'                db2.Execute "INSERT INTO `nonrcvd` SELECT * FROM " & dbase1 & ".`nonrcvd` "
'
'                db2.Execute "CREATE TABLE `grtrxfile` LIKE " & dbase1 & ".`grtrxfile`"
'                db2.Execute "INSERT INTO `grtrxfile` SELECT * FROM " & dbase1 & ".`grtrxfile` "
'
'                db2.Execute "CREATE TABLE `STAFFPYMT` LIKE " & dbase1 & ".`STAFFPYMT`"
'                db2.Execute "INSERT INTO `STAFFPYMT` SELECT * FROM " & dbase1 & ".`STAFFPYMT` "
'
'                db2.Execute "CREATE TABLE `Veh_Master` LIKE " & dbase1 & ".`Veh_Master`"
'                db2.Execute "INSERT INTO `Veh_Master` SELECT * FROM " & dbase1 & ".`Veh_Master` "
'
'                db2.Execute "CREATE TABLE `DAMAGE_MAST` LIKE " & dbase1 & ".`DAMAGE_MAST`"
'                db2.Execute "INSERT INTO `DAMAGE_MAST` SELECT * FROM " & dbase1 & ".`DAMAGE_MAST` "
'
'                db2.Execute "CREATE TABLE `DAMAGEMAST` LIKE " & dbase1 & ".`DAMAGEMAST`"
'                db2.Execute "INSERT INTO `DAMAGEMAST` SELECT * FROM " & dbase1 & ".`DAMAGEMAST` "
'
'                db2.Execute "CREATE TABLE `DGTRXFILEVAN` LIKE " & dbase1 & ".`DGTRXFILEVAN`"
'                db2.Execute "INSERT INTO `DGTRXFILEVAN` SELECT * FROM " & dbase1 & ".`DGTRXFILEVAN` "
'
'                db2.Execute "CREATE TABLE `DAMGMASTVAN` LIKE " & dbase1 & ".`DAMGMASTVAN`"
'                db2.Execute "INSERT INTO `DAMGMASTVAN` SELECT * FROM " & dbase1 & ".`DAMGMASTVAN` "
'
'                db2.Execute "CREATE TABLE `itemmastvan` LIKE " & dbase1 & ".`itemmastvan`"
'                db2.Execute "INSERT INTO `itemmastvan` SELECT * FROM " & dbase1 & ".`itemmastvan` "
'
'                db2.Execute "CREATE TABLE `transmastvan` LIKE " & dbase1 & ".`transmastvan`"
'                db2.Execute "INSERT INTO `transmastvan` SELECT * FROM " & dbase1 & ".`transmastvan` "
'
'                db2.Execute "CREATE TABLE `rtrxfilevan` LIKE " & dbase1 & ".`transmastvan`"
'                db2.Execute "INSERT INTO `rtrxfilevan` SELECT * FROM " & dbase1 & ".`transmastvan` "
'
'                db2.Execute "CREATE TABLE `trxfilevan` LIKE " & dbase1 & ".`trxfilevan`"
'                db2.Execute "INSERT INTO `trxfilevan` SELECT * FROM " & dbase1 & ".`trxfilevan` "
'
'                db2.Execute "CREATE TABLE `trxmastvan` LIKE " & dbase1 & ".`trxmastvan`"
'                db2.Execute "INSERT INTO `trxmastvan` SELECT * FROM " & dbase1 & ".`trxmastvan` "
'
'                db2.Execute "CREATE TABLE `trxsubvan` LIKE " & dbase1 & ".`trxsubvan`"
'                db2.Execute "INSERT INTO `trxsubvan` SELECT * FROM " & dbase1 & ".`trxsubvan` "
'
'
'                'day_book
'                '=======
'                db2.Execute "CREATE TABLE `prodlink` LIKE " & dbase1 & ".`prodlink`"
'                db2.Execute "INSERT INTO `prodlink` SELECT * FROM " & dbase1 & ".`prodlink`"
'
'                db2.Execute "CREATE TABLE `trxformulasub` LIKE " & dbase1 & ".`trxformulasub`"
'                db2.Execute "INSERT INTO `trxformulasub` SELECT * FROM " & dbase1 & ".`trxformulasub`"
'
'                db2.Execute "CREATE TABLE `trxformulamast` LIKE " & dbase1 & ".`trxformulamast`"
'                db2.Execute "INSERT INTO `trxformulamast` SELECT * FROM " & dbase1 & ".`trxformulamast`"
'                '========
'
'                '===============
'                db2.Execute "CREATE TABLE `docmast` LIKE " & dbase1 & ".`docmast`"
'                db2.Execute "INSERT INTO `docmast` SELECT * FROM " & dbase1 & ".`docmast`"
'
'                db2.Execute "CREATE TABLE `moleculelink` LIKE " & dbase1 & ".`moleculelink`"
'                db2.Execute "INSERT INTO `moleculelink` SELECT * FROM " & dbase1 & ".`moleculelink`"
'
'                db2.Execute "CREATE TABLE `molecules` LIKE " & dbase1 & ".`molecules`"
'                db2.Execute "INSERT INTO `molecules` SELECT * FROM " & dbase1 & ".`molecules`"
'
'                db2.Execute "CREATE TABLE `ptnmast` LIKE " & dbase1 & ".`ptnmast`"
'                db2.Execute "INSERT INTO `ptnmast` SELECT * FROM " & dbase1 & ".`ptnmast`"
'
'                db2.Execute "CREATE TABLE `SCHEDULE` LIKE " & dbase1 & ".`SCHEDULE`"
'                db2.Execute "INSERT INTO `SCHEDULE` SELECT * FROM " & dbase1 & ".`SCHEDULE`"
'                '=========
'
'                db2.Execute "CREATE TABLE `barprint` LIKE " & dbase1 & ".`barprint`"
'                db2.Execute "CREATE TABLE `ordermast` LIKE " & dbase1 & ".`ordermast`"
'                db2.Execute "CREATE TABLE `ordertrxfile` LIKE " & dbase1 & ".`ordertrxfile`"
'                db2.Execute "CREATE TABLE `roomtrxfile` LIKE " & dbase1 & ".`roomtrxfile`"
'                db2.Execute "CREATE TABLE `tbletrxfile` LIKE " & dbase1 & ".`tbletrxfile`"
'                db2.Execute "CREATE TABLE `trxfile_formula` LIKE " & dbase1 & ".`trxfile_formula`"
'                db2.Execute "CREATE TABLE `trnxroom` LIKE " & dbase1 & ".`trnxroom`"
'                db2.Execute "CREATE TABLE `trnxtable` LIKE " & dbase1 & ".`trnxtable`"
                              
                err.Clear
                On Error GoTo ERRHAND
                
                db2.Execute "CHECK TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
                db2.Execute "CHECK TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
                db2.Execute "CHECK TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
                db2.Execute "CHECK TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
                db2.Execute "CHECK TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
                db2.Execute "CHECK TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
                db2.Execute "CHECK TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
                db2.Execute "CHECK TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
                db2.Execute "CHECK TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
                db2.Execute "CHECK TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
                
                db2.Execute "OPTIMIZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
                db2.Execute "OPTIMIZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
                db2.Execute "OPTIMIZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
                db2.Execute "OPTIMIZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
                db2.Execute "OPTIMIZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
                db2.Execute "OPTIMIZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
                db2.Execute "OPTIMIZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
                db2.Execute "OPTIMIZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
                db2.Execute "OPTIMIZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
                db2.Execute "OPTIMIZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
                
                db2.Execute "REPAIR TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
                db2.Execute "REPAIR TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
                db2.Execute "REPAIR TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
                db2.Execute "REPAIR TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
                db2.Execute "REPAIR TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
                db2.Execute "REPAIR TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
                db2.Execute "REPAIR TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
                db2.Execute "REPAIR TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
                db2.Execute "REPAIR TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
                db2.Execute "REPAIR TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
                
                db2.Execute "ANALYZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
                db2.Execute "ANALYZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
                db2.Execute "ANALYZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
                db2.Execute "ANALYZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
                db2.Execute "ANALYZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
                db2.Execute "ANALYZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
                db2.Execute "ANALYZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
                db2.Execute "ANALYZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
                db2.Execute "ANALYZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
                db2.Execute "ANALYZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
                
                db2.Execute "FLUSH TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
                db2.Execute "FLUSH TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
                db2.Execute "FLUSH TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
                db2.Execute "FLUSH TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
                db2.Execute "FLUSH TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
                db2.Execute "FLUSH TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
                db2.Execute "FLUSH TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
                db2.Execute "FLUSH TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
                db2.Execute "FLUSH TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
                db2.Execute "FLUSH TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
                
                On Error Resume Next
                db2.Execute "CHECK TABLE `astmast`, `astrxfile`, `astrxmast` "
                db2.Execute "OPTIMIZE TABLE `astmast`, `astrxfile`, `astrxmast` "
                db2.Execute "REPAIR TABLE `astmast`, `astrxfile`, `astrxmast` "
                
                db2.Execute "CHECK TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
                db2.Execute "OPTIMIZE TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
                db2.Execute "REPAIR TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
                
                err.Clear
                On Error GoTo ERRHAND
                
       
                db2.Close
                Set db2 = Nothing
            
                
                Screen.MousePointer = vbHourglass
                strBackupEXT = "bktmp" & Format(Format(Date, "ddmmyy"), "000000") & Format(Format(Time, "HHMMSS"), "")
                DoEvents
                'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " > " & App.Path & "\Backup\" & strBackupEXT
                cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments tempdb > " & App.Path & "\Backup\" & strBackupEXT
                Call execCommand(cmd)
            End If
            
            
            Screen.MousePointer = vbNormal
            err.Clear
    End Select
    db.Execute "Update RTRXFILE set BAL_QTY = ROUND(BAL_QTY,2) where  BAL_QTY <>0"
    db.Execute "delete From TEMPTRXFILE "
    On Error Resume Next
    frmLogin.rs!LAST_LOGOUT = Format(Date, "DD/MM/YYYY") & " " & Format(Time, "hh:mm:ss")
    frmLogin.rs.Update
    frmLogin.rs.Close
    Set frmLogin.rs = Nothing
    db.Close
    Set db = Nothing
    
'    dbprint.Close
'    Set dbprint = Nothing
    Screen.MousePointer = vbNormal
    End
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub mnotkorder_Click()
    FrmTKOrder.Show
    FrmTKOrder.SetFocus
End Sub

Private Sub mnu2_Click()
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS2.Show
            FRMPOS2.SetFocus
        Else
            FRMSALES1.Show
            FRMSALES1.SetFocus
        End If
    Else
        If frmLogin.rs!Level = "5" Then
            FRMPOS2.Show
            FRMPOS2.SetFocus
        Else
            If SALESLT_FLAG = "Y" Then
                FRMGSTRSM2.Show
                FRMGSTRSM2.SetFocus
            Else
                FRMGSTR1.Show
                FRMGSTR1.SetFocus
            End If
        End If
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
End Sub

Private Sub mnuabstract_Click()
    FRMABSTRACT.Show
    FRMABSTRACT.SetFocus
End Sub

Private Sub mnuabt_Click()
    frmabout.Show
End Sub

Private Sub MNUAGNT_Click()
    frmAgentMast.Show
    frmAgentMast.SetFocus
End Sub

Private Sub mnuAMC_Click()
    FrmAMC.Show
    FrmAMC.SetFocus
End Sub

Private Sub mnuarea_Click()
    FRMAREAWISERPT.Show
    FRMAREAWISERPT.SetFocus
End Sub

Private Sub mnuassets_Click()
    FRMAsstReg.Show
    FRMAsstReg.SetFocus
End Sub

Private Sub MNUASTMASTR_Click()
    frmAstmaster.Show
    frmAstmaster.SetFocus
End Sub

Private Sub mnuAstPurchase_Click()
    frmAstPurchase.Show
    frmAstPurchase.SetFocus
End Sub

Private Sub mnub2c1_Click()
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS1.Show
            FRMPOS1.SetFocus
        Else
            frmsales.Show
            frmsales.SetFocus
        End If
    Else
        If frmLogin.rs!Level = "5" Then
            FRMPOS1.Show
            FRMPOS1.SetFocus
        Else
            If SALESLT_FLAG = "Y" Then
                FRMGSTRSM1.Show
                FRMGSTRSM1.SetFocus
            Else
                FRMGSTR.Show
                FRMGSTR.SetFocus
            End If
        End If
    End If
End Sub

Private Sub MNUBACK_Click()
    'Exit Sub
    fRMbackup.Show
    fRMbackup.SetFocus
End Sub

Private Sub MNUBANKBOOK_Click()
    FRMBankBook.Show
    FRMBankBook.SetFocus
End Sub

Private Sub MNUBANKMSTR_Click()
    frmbankMast.Show
    frmbankMast.SetFocus
End Sub

Private Sub mnubr_Click()
    FRMDAMAGEREGBR.Show
    FRMDAMAGEREGBR.SetFocus
End Sub

Private Sub mnubrmaster_Click()
    frmbranchmast.Show
    frmbranchmast.SetFocus
End Sub

Private Sub mnuBrsalereg_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMBILLPRINTBR.Show
    FRMBILLPRINTBR.SetFocus
End Sub

Private Sub mnucalc_Click()
    frmCalculator1.Show
    frmCalculator1.SetFocus
    frmCalculator1.WindowState = Normal
End Sub

Private Sub MNUCASHBOOK_Click()
    FRMCASHBOOK.Show
    FRMCASHBOOK.SetFocus
End Sub

Private Sub mnucat_Click()
    FrmCatmaster.Show
    FrmCatmaster.SetFocus
End Sub

Private Sub MNUCOMSN_Click()
    FRMAGENTREG.Show
    FRMAGENTREG.SetFocus
End Sub

Private Sub mnucontra_Click()
    FRMContra.Show
    FRMContra.SetFocus
End Sub

Private Sub MNUCOST_Click()
    Dim RSTCOST As ADODB.Recordset
    Dim rstTRANX As ADODB.Recordset
    Dim GROSSAMT As Double
    Dim COSTAMT As Double
    
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From TRXMAST WHERE (TRX_TYPE='GI' OR TRX_TYPE='HI' OR TRX_TYPE='SV') ORDER BY VCH_NO", db, adOpenStatic, adLockOptimistic, adCmdText
    rstTRANX.Properties("Update Criteria").Value = adCriteriaKey
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    Do Until rstTRANX.EOF
        GROSSAMT = 0
        COSTAMT = 0
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT * From TRXFILE WHERE TRX_TYPE= '" & rstTRANX!TRX_TYPE & "' AND TRX_YEAR= '" & rstTRANX!TRX_YEAR & "' AND VCH_NO= " & rstTRANX!VCH_NO & "  ", db, adOpenForwardOnly
        Do Until RSTCOST.EOF
            GROSSAMT = Round(GROSSAMT + (IIf(IsNull(RSTCOST!PTR), 0, RSTCOST!PTR) * IIf(IsNull(RSTCOST!QTY), 0, RSTCOST!QTY)) - (IIf(IsNull(RSTCOST!PTR), 0, RSTCOST!PTR) * IIf(IsNull(RSTCOST!QTY), 0, RSTCOST!QTY)) * IIf(IsNull(RSTCOST!LINE_DISC), 0, RSTCOST!LINE_DISC) / 100, 2)
            COSTAMT = COSTAMT + IIf(IsNull(RSTCOST!item_COST), 0, RSTCOST!item_COST) * (IIf(IsNull(RSTCOST!QTY), 0, RSTCOST!QTY) + IIf(IsNull(RSTCOST!FREE_QTY), 0, RSTCOST!FREE_QTY))
            RSTCOST.MoveNext
        Loop
        RSTCOST.Close
        Set RSTCOST = Nothing
        'IIF(ISNULL(rstTRANX!LINE_DISC),0,rstTRANX!LINE_DISC)
        
        
        rstTRANX!gross_amt = GROSSAMT
        rstTRANX!PAY_AMOUNT = COSTAMT
        rstTRANX.Update
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From ITEMMAST ORDER BY ITEM_CODE", db, adOpenStatic, adLockOptimistic, adCmdText
    rstTRANX.Properties("Update Criteria").Value = adCriteriaKey
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    Do Until rstTRANX.EOF
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT * From  RTRXFILE WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' AND ITEM_COST <>0 AND NOT ISNULL(ITEM_COST) ORDER BY VCH_DATE DESC  ", db, adOpenForwardOnly
        If Not (RSTCOST.EOF And RSTCOST.BOF) Then
            rstTRANX!item_COST = RSTCOST!item_COST
            rstTRANX!SALES_TAX = RSTCOST!SALES_TAX
            rstTRANX.Update
        End If
        RSTCOST.Close
        
        Set RSTCOST = Nothing
        
        Set RSTCOST = New ADODB.Recordset
        RSTCOST.Open "SELECT * From  RTRXFILE WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' AND P_RETAIL <>0 AND NOT ISNULL(P_RETAIL) ORDER BY VCH_DATE DESC  ", db, adOpenForwardOnly
        If Not (RSTCOST.EOF And RSTCOST.BOF) Then
            rstTRANX!P_RETAIL = RSTCOST!P_RETAIL
            rstTRANX!MRP = RSTCOST!MRP
            rstTRANX.Update
        End If
        RSTCOST.Close
        Set RSTCOST = Nothing
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuCounter_Click()
    FRMCounterReg.Show
    FRMCounterReg.SetFocus
End Sub

Private Sub MNUCUST_Click()
    frmcustmast1.Show
    frmcustmast1.SetFocus
End Sub

Private Sub MNUCUSTLIST_Click()
    FrmCustList.Show
    FrmCustList.SetFocus
End Sub

Private Sub MNUDAMAGE_Click()
    FRMDAMAGE.Show
    FRMDAMAGE.SetFocus
End Sub

Private Sub mnudamge_Click()
    frmDmgRet.Show
    frmDmgRet.SetFocus
End Sub

Private Sub MNUDAMREG_Click()
    FRMDAMAGEREG.Show
    FRMDAMAGEREG.SetFocus
End Sub

Private Sub mnuDead_Click()
    FrmDead.Show
    FrmDead.SetFocus
End Sub

Private Sub MNUDEL_REG_Click()
    FRMDELREG.Show
    FRMDELREG.SetFocus
End Sub

Private Sub mnudelitems_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    frmdelitems.Show
    frmdelitems.SetFocus
End Sub

Private Sub mnudelzero_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim rststock As ADODB.Recordset
   
    Dim i As Long
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * FROM ITEMMAST ", db, adOpenForwardOnly
    vbalProgressBar1.Min = 0
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT DISTINCT ITEM_CODE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'", db, adOpenForwardOnly
        If rststock.RecordCount > 0 Then
            rststock.Close
            Set rststock = Nothing
            GoTo SKIP
        End If
        rststock.Close
        Set rststock = Nothing
            
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT DISTINCT ITEM_CODE from TRXFILE where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'", db, adOpenForwardOnly
        If rststock.RecordCount > 0 Then
            rststock.Close
            Set rststock = Nothing
            GoTo SKIP
        End If
        rststock.Close
        Set rststock = Nothing
        
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT DISTINCT FOR_NAME from TRXFORMULASUB where FOR_NAME = '" & rstTRANX!ITEM_CODE & "'", db, adOpenForwardOnly
        If rststock.RecordCount > 0 Then
            rststock.Close
            Set rststock = Nothing
            GoTo SKIP
        End If
        rststock.Close
        Set rststock = Nothing

        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT DISTINCT ITEM_CODE from TRXFORMULAMAST where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'", db, adOpenForwardOnly
        If rststock.RecordCount > 0 Then
            rststock.Close
            Set rststock = Nothing
            GoTo SKIP
        End If
        rststock.Close
        Set rststock = Nothing
        
        
        'db.Execute ("DELETE from RTRXFILE where RTRXFILE.ITEM_CODE = '" & TXTITEMCODE.Text & "'")
        db.Execute ("DELETE from PRODLINK where PRODLINK.ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'")
        db.Execute ("DELETE from ITEMMAST where ITEMMAST.ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'")
        
        i = i + 1
SKIP:
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox i & " ITEMS", , "EzBiz"
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    vbalProgressBar1.Visible = False
    Exit Sub
ERRHAND:
    vbalProgressBar1.Value = 0
    vbalProgressBar1.Text = ""
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnudeposit_Click()
    frmDeposit.Show
    frmDeposit.SetFocus
End Sub

Private Sub mnudepositentry_Click()
    FRMDPReg.Show
    FRMDPReg.SetFocus
End Sub

Private Sub mnudmgbr_Click()
    FRMDAMAGEBR.Show
    FRMDAMAGEBR.SetFocus
End Sub

Private Sub mnudncnreg_Click()
    FRMDNCNREG.Show
    FRMDNCNREG.SetFocus
End Sub

Private Sub MNUEMPLOYEE_Click()
    frmEmployeeMast.Show
    frmEmployeeMast.SetFocus
End Sub

Private Sub mnuexit_Click()
    db.Close
    Set db = Nothing
    
'    dbprint.Close
'    Set dbprint = Nothing
    End
    
    'Unload Me
End Sub

Private Sub MnuExpenseSt_Click()
    If MDIMAIN.lblsalary.Caption = "Y" Then
        FRMStaffReg.Show
        FRMStaffReg.SetFocus
    Else
        frmExpenseStaff.Show
        frmExpenseStaff.SetFocus
    End If
End Sub

Private Sub MNUEXPENTRY_Click()
    'If PCTMENU.Visible = True Then
        Frmexpense.Show
        Frmexpense.SetFocus
    'Else
        'frmExpensewo.Show
        'frmExpensewo.SetFocus
    'End If
End Sub

Private Sub MNUEXPLDGR_Click()
    frmExpmast.Show
    frmExpmast.SetFocus
End Sub

Private Sub mnuexportMc_Click()
    fRMPluUpdate.Show
    fRMPluUpdate.SetFocus
End Sub

Private Sub mnuExpPurchase_Click()
    frmExpPurchase.Show
    frmExpPurchase.SetFocus
End Sub

Private Sub MnuExpReg_Click()
    FrmExpReg.Show
    FrmExpReg.SetFocus
End Sub

Private Sub MnuExpStaff_Click()
    frmExpenseStaff.Show
    frmExpenseStaff.SetFocus
End Sub

Private Sub mnuExpRegtax_Click()
    FRMEXPINREGISTER.Show
    FRMEXPINREGISTER.SetFocus
End Sub

Private Sub mnuexpret_Click()
    frmExpiryRet.Show
    frmExpiryRet.SetFocus
End Sub

Private Sub MnuExpStaffReg_Click()
    FrmExpStaffReg.Show
    FrmExpStaffReg.SetFocus
End Sub

Private Sub mnuezbill_Click()
    frmeasybill.Show
    frmeasybill.SetFocus
End Sub

Private Sub mnuFAReg_Click()
    FrmFAReg.Show
    FrmFAReg.SetFocus
End Sub

Private Sub mnufix_Click()
    If MsgBox("Are you sure?", vbYesNo, "Fix Error") = vbNo Then Exit Sub
    On Error GoTo ERRHAND
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

Private Sub MNUFMLA_Click()
    FRMFormula.Show
    FRMFormula.SetFocus
End Sub

Private Sub MNUFREE_Click()
    FRMFREE.Show
    FRMFREE.SetFocus
End Sub

Private Sub mnuFxdAssets_Click()
    frmFixedAssets.Show
    frmFixedAssets.SetFocus
End Sub

Private Sub mnugift_Click()
    FRMSAMPLEGOODS.Show
    FRMSAMPLEGOODS.SetFocus
End Sub

Private Sub MNUGIFTREG_Click()
'    If PCTMENU.Visible = True Then
        FRMGIFTREG.Show
        FRMGIFTREG.SetFocus
'    End If
End Sub

Private Sub MNUGSTBTBU_Click()
     If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        
    Else
        FRMGSTUN.Show
        FRMGSTUN.SetFocus
    End If
End Sub

Private Sub MNUGSTSERVICE_Click()
    FRMGSTSERVICEUN.Show
    FRMGSTSERVICEUN.SetFocus
End Sub

Private Sub mnuimport_Click()
    frmOnline.Show
    frmOnline.SetFocus
End Sub

Private Sub mnuincom_Click()
    frmIncome.Show
    frmIncome.SetFocus
End Sub

Private Sub mnuincome_Click()
    FrmIncReg.Show
    FrmIncReg.SetFocus
End Sub

Private Sub mnuinout_Click()
    frmstkmvmreport.Show
    frmstkmvmreport.SetFocus
End Sub

Private Sub mnuinoutbr_Click()
    frmstkmvmreportvan.Show
    frmstkmvmreportvan.SetFocus
End Sub

Private Sub MNUIPCOST_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim GROSSAMT As Double
    Dim COSTAMT As Double
    
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * From ITEMMAST ORDER BY ITEM_CODE", db, adOpenForwardOnly
    rstTRANX.Properties("Update Criteria").Value = adCriteriaKey
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    Do Until rstTRANX.EOF
        db.Execute "Update RTRXFILE set GROSS_AMT = (QTY * " & rstTRANX!item_COST & "), TRX_TOTAL = (QTY * " & rstTRANX!item_COST & ")+((QTY * " & rstTRANX!item_COST & ") * SALES_TAX / 100), ITEM_COST = " & rstTRANX!item_COST & ", PTR = " & rstTRANX!item_COST & " where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' AND BAL_QTY >0 AND (TRX_TYPE = 'OP' OR TRX_TYPE = 'ST')"
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuitemmast_Click()
    frmitemmaster.Show
    frmitemmaster.SetFocus
End Sub

Private Sub mnujournal_Click()
    FRMJournal.Show
    FRMJournal.SetFocus
End Sub

Private Sub mnulabel_Click()
    FrmBarcodePrint.Show
    FrmBarcodePrint.SetFocus
End Sub

Private Sub MNULITEMLST_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT ITEM_CODE FROM RTRXFILE ", db, adOpenForwardOnly
    vbalProgressBar1.Min = 0
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from ITEMMAST where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "'  ", db, adOpenStatic, adLockOptimistic, adCmdText
        If (rststock.EOF And rststock.BOF) Then
            Set RSTITEM = New ADODB.Recordset
            RSTITEM.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
            If Not (RSTITEM.EOF And RSTITEM.BOF) Then
                i = i + 1
                rststock.AddNew
                rststock!ITEM_CODE = RSTITEM!ITEM_CODE
                rststock!ITEM_NAME = RSTITEM!ITEM_NAME
                rststock!Category = RSTITEM!Category
                rststock!UNIT = 1
                rststock!MANUFACTURER = RSTITEM!MFGR
                rststock!DEAD_STOCK = "N"
                rststock!UN_BILL = "N"
                rststock!PRICE_CHANGE = "N"
                rststock!REMARKS = ""
                rststock!ITEM_SPEC = ""
                rststock!REORDER_QTY = 1
                rststock!PACK_TYPE = RSTITEM!PACK_TYPE
                rststock!FULL_PACK = RSTITEM!PACK_TYPE
                rststock!BIN_LOCATION = ""
                rststock!ITEM_MAL = ""
                rststock!PTR = 0
                rststock!CST = 0
                rststock!OPEN_QTY = 0
                rststock!OPEN_VAL = 0
                rststock!RCPT_QTY = 0
                rststock!RCPT_VAL = 0
                rststock!ISSUE_QTY = 0
                rststock!ISSUE_VAL = 0
                rststock!CLOSE_QTY = 0
                rststock!CLOSE_VAL = 0
                rststock!DAM_QTY = 0
                rststock!DAM_VAL = 0
                rststock!DISC = 0
                rststock!SALES_TAX = RSTITEM!SALES_TAX
                rststock!check_flag = "V"
                rststock!item_COST = RSTITEM!item_COST
                rststock!P_RETAIL = RSTITEM!P_RETAIL
                rststock!MRP = RSTITEM!MRP
                rststock!P_WS = RSTITEM!P_WS
                rststock!CRTN_PACK = RSTITEM!CRTN_PACK
                rststock!P_CRTN = RSTITEM!P_CRTN
                rststock!LOOSE_PACK = RSTITEM!LOOSE_PACK
'                rststock!PACK_DESC = rstITEM!ITEM_COST
'                rststock!PACK_DET = Val(txtpackdet.Text)
                rststock!BARCODE = RSTITEM!BARCODE
        
                rststock.Update
            End If
            RSTITEM.Close
            Set RSTITEM = Nothing
        Else
            If IsNull(rststock!ITEM_NAME) Or rststock!ITEM_NAME = "" Or rststock!ITEM_NAME = "." Then
                Set RSTITEM = New ADODB.Recordset
                RSTITEM.Open "SELECT * FROM RTRXFILE where ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenForwardOnly
                If Not (RSTITEM.EOF And RSTITEM.BOF) Then
                    rststock!ITEM_NAME = RSTITEM!ITEM_NAME
                    rststock.Update
                End If
                RSTITEM.Close
                Set RSTITEM = Nothing
            End If
            
        End If
        rststock.Close
        Set rststock = Nothing
        
        'db.Execute "Update RTRXFILE SET P_RETAIL = " & rstTRANX!P_RETAIL & " WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' "
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox i & " ITEMS", , "EzBiz"
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    vbalProgressBar1.Value = 0
    vbalProgressBar1.Text = ""
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub MnuLoc_Click()
    'If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
    FrmLocSearch.Show
    FrmLocSearch.SetFocus
End Sub

Private Sub mnumerge_Click()
    Frmitemmerge.Show
    Frmitemmerge.SetFocus
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
    
'
'    Dim RSTRTRXFILE As ADODB.Recordset
'    Dim rststock As ADODB.Recordset
'    Dim rstTRXMAST As ADODB.Recordset
'    Dim I As Long
'
'    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
'    On Error GoTo ErrHand
'    Screen.MousePointer = vbHourglass
'    vbalProgressBar1.Visible = True
'    vbalProgressBar1.value = 0
'    vbalProgressBar1.ShowText = True
'    vbalProgressBar1.Text = "PLEASE WAIT..."
'    Screen.MousePointer = vbHourglass
'    Set RSTRTRXFILE = New ADODB.Recordset
'    RSTRTRXFILE.Open "SELECT DISTINCT ITEM_CODE FROM RTRXFILE WHERE  RTRXFILE.BAL_QTY > 0", db, adOpenForwardOnly
'    Do Until RSTRTRXFILE.EOF
'        Set rststock = New ADODB.Recordset
'        rststock.Open "SELECT * from RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & RSTRTRXFILE!ITEM_CODE & "'  AND RTRXFILE.BAL_QTY > 0 ", db, adOpenForwardOnly
'        Do Until rststock.EOF
'            I = 0
'            Set rstTRXMAST = New ADODB.Recordset
'            rstTRXMAST.Open "SELECT * from RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & rststock!ITEM_CODE & "' AND RTRXFILE.P_RETAIL = " & rststock!P_RETAIL & " AND RTRXFILE.BAL_QTY > 0 ORDER BY RTRXFILE.VCH_NO", db, adOpenStatic, adLockOptimistic, adCmdText
'            'EXP_DATE <=# " & E_DATE & " #
'            If rstTRXMAST.RecordCount > 1 Then
'                Do Until rstTRXMAST.EOF
'                    I = I + rstTRXMAST!BAL_QTY
'                    rstTRXMAST!BAL_QTY = 0
'                    rstTRXMAST.Update
'                    rstTRXMAST.MoveNext
'                Loop
'                rstTRXMAST.MoveLast
'                rstTRXMAST!BAL_QTY = I
'                rstTRXMAST.Update
'            End If
'            rstTRXMAST.Close
'            Set rstTRXMAST = Nothing
'            rststock.MoveNext
'        Loop
'        rststock.Close
'        Set rststock = Nothing
'
'        Screen.MousePointer = vbNormal
'        vbalProgressBar1.Max = RSTRTRXFILE.RecordCount
'        vbalProgressBar1.value = vbalProgressBar1.value + 1
'        RSTRTRXFILE.MoveNext
'    Loop
'    RSTRTRXFILE.Close
'    Set RSTRTRXFILE = Nothing
'
'    vbalProgressBar1.Text = "Successfully Completed..."
'    MsgBox "Item Merge complete !!", vbOKOnly, "Item Merge!!!!"
'    vbalProgressBar1.Visible = False
'    Screen.MousePointer = vbNormal
'    Exit Sub
'ErrHand:
'    Screen.MousePointer = vbNormal
'    MsgBox Err.Description
End Sub

Private Sub mnumergemulti_Click()
    FrmitemmergeMulti.Show
    FrmitemmergeMulti.SetFocus
End Sub

Private Sub MnuMinQty_Click()
    FRMSETMINSTOCK.Show
    FRMSETMINSTOCK.SetFocus
End Sub

Private Sub mnumnth_Click()
    FrmSaleAnalysis.Show
    FrmSaleAnalysis.SetFocus
End Sub

Private Sub mnuOP_Click()
    If MDIMAIN.LBLSHOPRT.Caption = "Y" Then
        frmOPS.Show
        frmOPS.SetFocus
    Else
        frmOP.Show
        frmOP.SetFocus
    End If
End Sub

Private Sub mnuOPCash_Click()
    frmOPCash.Show
    frmOPCash.SetFocus
End Sub

Private Sub MNUOPSTK_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    frmOPSTK.Show
    frmOPSTK.SetFocus
End Sub

Private Sub mnuorder_Click()
    FrmOrder2.Show
    FrmOrder2.SetFocus
End Sub

Private Sub MNUP_RETURNREG_Click()
    FRMPR.Show
    FRMPR.SetFocus
End Sub

Private Sub mnuplist_Click()
    FrmBinLoc1.Show
    FrmBinLoc1.SetFocus
End Sub

Private Sub mnuPluItems_Click()
    frmplumaster.Show
    frmplumaster.SetFocus
End Sub

Private Sub MNUPO_Click()
    frmPO.Show
    frmPO.SetFocus
End Sub

Private Sub mnuprice_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FrmPriceAnalysis.Show
    FrmPriceAnalysis.SetFocus
End Sub

Private Sub mnupricelist_Click()
    FrmBinLoc.Show
    FrmBinLoc.SetFocus
End Sub

Private Sub MNUPRICEUPDATE_Click()
    
    Dim rstTRANX As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT * from ITEMMAST where PRICE_CHANGE = 'Y' ", db, adOpenForwardOnly
    vbalProgressBar1.Min = 0
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from RTRXFILE where ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' AND BAL_QTY >0 ", db, adOpenStatic, adLockOptimistic, adCmdText
        Do Until rststock.EOF
            rststock!P_RETAIL = IIf(IsNull(rstTRANX!P_RETAIL), 0, rstTRANX!P_RETAIL)
            rststock.Update
            rststock.MoveNext
        Loop
        rststock.Close
        Set rststock = Nothing
        
        'db.Execute "Update RTRXFILE SET P_RETAIL = " & rstTRANX!P_RETAIL & " WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' "
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    MDIMAIN.vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    MDIMAIN.vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.Text = ""
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub MNUPRINTMIX_Click()
    On Error GoTo ERRHAND
    Dim i As Integer
    ReportNameVar = Rptpath & "RptPRODUCTION"
    Set Report = crxApplication.OpenReport(ReportNameVar, 1)
'    If OPTCUSTOMER.value = True Then
'        Report.RecordSelectionFormula = "({DBTPYMT.ACT_CODE}='" & DataList2.BoundText & "' and ({DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} ='DR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
'    Else
'        Report.RecordSelectionFormula = "(({DBTPYMT.TRX_TYPE} ='RT' OR {DBTPYMT.TRX_TYPE} ='SR' OR {DBTPYMT.TRX_TYPE} ='DR') AND {DBTPYMT.INV_DATE} <=# " & Format(DTTO.value, "MM,DD,YYYY") & " # AND {DBTPYMT.INV_DATE} >=# " & Format(DTFROM.value, "MM,DD,YYYY") & " #)"
'    End If

    Set CRXFormulaFields = Report.FormulaFields
    For i = 1 To Report.Database.Tables.COUNT
        Report.Database.Tables.Item(i).SetLogOnInfo strConnection
        If UCase(dbase1) <> "INVSOFT" And UCase(dbase1) <> "INVDB" And UCase(dbase1) <> "INVSOFT3" Then
            Set oRs = New ADODB.Recordset
            Set oRs = db.Execute("SELECT * FROM " & Report.Database.Tables(i).Name & " ")
            Report.Database.Tables(i).SetDataSource oRs, 3
            Set oRs = Nothing
        End If
    Next i
    Report.DiscardSavedData
    Report.VerifyOnEveryPrint = True
    frmreport.Caption = "REPORT"
    Call GENERATEREPORT
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub MNUPROCESS_Click()
    FRMProcess.Show
    FRMProcess.SetFocus
End Sub

Private Sub mnuproduction_Click()
    FRMPRODUCTION.Show
    FRMPRODUCTION.SetFocus
End Sub

Private Sub mnupromotion_Click()
    FRMPROMOREG.Show
    FRMPROMOREG.SetFocus
End Sub

Private Sub MNUPROREG_Click()
    FRMPRODREP.Show
    FRMPRODREP.SetFocus
End Sub

Private Sub mnupurchase_Click()
    If cmdpurchase.Enabled = True Then cmdpurchase_Click
End Sub

Private Sub mnupymntdue_Click()
    FRMPYMNTDUES.Show
    FRMPYMNTDUES.SetFocus
End Sub

Private Sub MnuQtn_Click()
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From QTNMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    FRMQUOTATION.Show
    FRMQUOTATION.SetFocus
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description, , "EzBiz"
End Sub

Private Sub MNUQTNVIEW_Click()
    FRMQTNREG.Show
    FRMQTNREG.SetFocus
End Sub

Private Sub mnurcptdue_Click()
    FRMRCPTDUES.Show
    FRMRCPTDUES.SetFocus
End Sub

Private Sub MNURCPTREG_Click()
    FRMRcptReg.Show
    FRMRcptReg.SetFocus
End Sub

Private Sub mnurcptregstr_Click()
    On Error Resume Next
    FRMRecpt.Show
    FRMRecpt.SetFocus
End Sub

Private Sub mnurcvdist_Click()
    FRMEXPRCVD.Show
    FRMEXPRCVD.SetFocus
End Sub

Private Sub MNURCVGOODS_Click()
    FRMDef_Rcvd.Show
    FRMDef_Rcvd.SetFocus
End Sub

Private Sub MNUREFRESH_Click()
    Call CommandButton1_Click
End Sub

Private Sub mnuremind_Click()
    FrmOccassion.Show
    FrmOccassion.SetFocus
End Sub

Private Sub mnuReminder_Click()
    Frmreminder.Show
    Frmreminder.SetFocus
End Sub

Private Sub MNURESTRECUST_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Dim i As Long
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT ACT_CODE FROM TRXMAST ", db, adOpenForwardOnly
    vbalProgressBar1.Min = 0
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from CUSTMAST where CUST_CODE = '" & rstTRANX!ACT_CODE & "'  ", db, adOpenStatic, adLockOptimistic, adCmdText
        If (rststock.EOF And rststock.BOF) Then
            Set RSTITEM = New ADODB.Recordset
            RSTITEM.Open "SELECT * FROM TRXMAST where ACT_CODE = '" & rstTRANX!CUST_CODE & "' ", db, adOpenForwardOnly
            If Not (RSTITEM.EOF And RSTITEM.BOF) Then
                i = i + 1
                rststock.AddNew
                rststock!CUST_CODE = RSTITEM!ACT_CODE
                rststock!cust_name = RSTITEM!ACT_NAME
                 
                rststock!Address = ""
                rststock!TELNO = ""
                rststock!FAXNO = ""
                rststock!EMAIL_ADD = ""
                rststock!DL_NO = ""
                rststock!REMARKS = ""
                rststock!KGST = ""
                rststock!CST = ""
                rststock!PYMT_PERIOD = 0
                rststock!PYMT_LIMIT = 0
                rststock!Area = ""
                rststock!AGENT_CODE = ""
                rststock!AGENT_NAME = ""
                rststock!CONTACT_PERSON = "CS"
                rststock!SLSM_CODE = "SM"
                rststock!OPEN_DB = 0
                rststock!OPEN_CR = 0
                rststock!YTD_DB = 0
                rststock!YTD_CR = 0
                rststock!CREATE_DATE = Date
                rststock!C_USER_ID = "SM"
                rststock!MODIFY_DATE = Date
                rststock!M_USER_ID = "SM"
                rststock!Type = "R"
                rststock!CUST_TYPE = ""
                rststock!CUST_IGST = ""
                rststock.Update
            End If
            RSTITEM.Close
            Set RSTITEM = Nothing
        End If
        rststock.Close
        Set rststock = Nothing
        
        'db.Execute "Update RTRXFILE SET P_RETAIL = " & rstTRANX!P_RETAIL & " WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' "
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox i & " CUSTOMERS", , "EzBiz"
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    vbalProgressBar1.Value = 0
    vbalProgressBar1.Text = ""
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub mnuretail_Click()
    
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
'    On Error GoTo eRRhAND
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'        MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'        rstTRXMAST.Close
'        Set rstTRXMAST = Nothing
'        Exit Sub
'    End If
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS3.Show
            FRMPOS3.SetFocus
        Else
            FRMSALES2.Show
            FRMSALES2.SetFocus
        End If
    Else
        If frmLogin.rs!Level = "5" Then
            FRMPOS3.Show
            FRMPOS3.SetFocus
        Else
            If SALESLT_FLAG = "Y" Then
                FRMGSTRSM3.Show
                FRMGSTRSM3.SetFocus
            Else
                FRMGSTR2.Show
                FRMGSTR2.SetFocus
            End If
        End If
    End If
'    Exit Sub
'eRRhAND:
'    MsgBox Err.Description
End Sub

Private Sub mnureturn_Click()
    FRMRETURN.Show
    FRMRETURN.SetFocus
End Sub

Private Sub mnurordr_Click()
    FRMReorderStk.Show
    FRMReorderStk.SetFocus
End Sub

Private Sub MNUROUTE_Click()
    FRMROUTE.Show
    FRMROUTE.SetFocus
End Sub

Private Sub MNUS_RETURNREG_Click()
    FRMSALEREG.Show
    FRMSALEREG.SetFocus
End Sub

Private Sub mnusalesrep_Click()
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    db.Execute "Update ITEMMAST set UQC = 'NOS' where pack_type = 'Nos' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'KGS' where pack_type = 'Kg' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'BAG' where pack_type = 'Bag' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'BOX' where pack_type = 'Box' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'DOZ' where pack_type = 'doz' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'GMS' where pack_type = 'gm' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'PCS' where pack_type = 'Pcs' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'PRS' where pack_type = 'Pair' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'PAC' where pack_type = 'Pkt' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'BDL' where pack_type = 'Bndl' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'PAC' where pack_type = 'Pkt' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'MTR' where pack_type = 'Meters' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'TON' where pack_type = 'Ton' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'SQF' where pack_type = 'sqft' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'SQM' where pack_type = 'sqmtr' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'ROL' where pack_type = 'Roll' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'MLT' where pack_type = 'ml' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'CBM' where pack_type = 'cbm' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'CTN' where pack_type = 'Cases' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'OTH' where pack_type = 'cu ft' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'OTH' where pack_type = 'each' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'OTH' where pack_type = 'ft' And isnull(UQC)"
    db.Execute "Update ITEMMAST set UQC = 'OTH' where pack_type = 'Litres' And isnull(UQC)"
    
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMSalesReg.Show
    FRMSalesReg.SetFocus
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub mnusentgoods_Click()
    FRMWARRANTY.Show
    FRMWARRANTY.SetFocus
End Sub

Private Sub MnuSer_Stk_Click()
    FRMServiceStk.Show
    FRMServiceStk.SetFocus
End Sub

Private Sub MnuSerStkReg_Click()
    FRMServReg.Show
    FRMServReg.SetFocus
End Sub

Private Sub MNUSERVICES_Click()
    FrmSrvmovmnt.Show
    FrmSrvmovmnt.SetFocus
End Sub

Private Sub MNUSHOPINFO_Click()
    FRMSHOINFO.Show
    FRMSHOINFO.SetFocus
End Sub

Private Sub mnuStk_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMSTKSUMMRY.Show
    FRMSTKSUMMRY.SetFocus
End Sub

Private Sub MNUSTKANALYSIS_Click()
    FRMSTOCK.Show
    FRMSTOCK.SetFocus
End Sub

Private Sub mnustkcr_Click()
    If Not (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then Exit Sub
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FrmStockCorrect.Show
    FrmStockCorrect.SetFocus
End Sub

Private Sub mnustkmv_Click()
    FrmStkmovmntvan.Show
    FrmStkmovmntvan.SetFocus
End Sub

Private Sub mnustksum_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMSTKSUMRy.Show
    FRMSTKSUMRy.SetFocus
End Sub

Private Sub mnustktransfer_Click()
    frmSTKTRANSFER.Show
    frmSTKTRANSFER.SetFocus
End Sub

Private Sub MNUSUPPLIER_Click()
    frmsuppliermast.Show
    frmsuppliermast.SetFocus
End Sub

Private Sub mnusync_Click()
    On Error GoTo ERRHAND
    Screen.MousePointer = vbHourglass
    If REMOTEAPP = 1 Then
        Call export_db
    Else
        Call export_db2
    End If
    
    Screen.MousePointer = vbNormal
    Exit Sub
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Sub

Private Sub mnutable_Click()
    frmTablemast.Show
    frmTablemast.SetFocus
End Sub

Private Sub mnutblewise_Click()
    FRMTO.Show
    FRMTO.SetFocus
End Sub

Private Sub mnutest_Click()
    Dim RSTTRXFILE As ADODB.Recordset
    Dim RSTTRXFILE2 As ADODB.Recordset
    Dim RSTTRXFILE3 As ADODB.Recordset
    Dim a, c, E As Long
    Dim b, D, F As Double
    Set RSTTRXFILE = New ADODB.Recordset
    RSTTRXFILE.Open "SELECT *  FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
    With RSTTRXFILE
        Do Until .EOF
            a = 0
            b = 0
            c = 0
            D = 0
            E = 0
            F = 0
            Set RSTTRXFILE2 = New ADODB.Recordset
            RSTTRXFILE2.Open "SELECT *  FROM RTRXFILE WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE2.EOF
                a = a + RSTTRXFILE2!QTY                                 'Rcpt qty
                b = b + RSTTRXFILE2!TRX_TOTAL                           'Rcpt Value
                c = c + RSTTRXFILE2!BAL_QTY                             'Closing Qty
                D = D + RSTTRXFILE2!item_COST * RSTTRXFILE2!BAL_QTY     'Closing Value
                RSTTRXFILE2.MoveNext
            Loop
            RSTTRXFILE2.Close
            Set RSTTRXFILE2 = Nothing
            
            Set RSTTRXFILE3 = New ADODB.Recordset
            RSTTRXFILE3.Open "SELECT *  FROM TRXFILE WHERE ITEM_CODE = '" & RSTTRXFILE!ITEM_CODE & "'", db, adOpenStatic, adLockReadOnly, adCmdText
            Do Until RSTTRXFILE3.EOF
                E = E + RSTTRXFILE3!QTY + RSTTRXFILE3!FREE_QTY  'Issue qty
                F = F + RSTTRXFILE3!TRX_TOTAL                   'Issue Value
                RSTTRXFILE3.MoveNext
            Loop
            RSTTRXFILE3.Close
            Set RSTTRXFILE3 = Nothing
            
            !RCPT_QTY = a
            !RCPT_VAL = b
            !CLOSE_QTY = c
            !CLOSE_VAL = D
            !ISSUE_QTY = E
            !ISSUE_VAL = F
            RSTTRXFILE.Update
            RSTTRXFILE.MoveNext
        Loop
    End With
    RSTTRXFILE.Close
    Set RSTTRXFILE = Nothing
End Sub

Private Sub mnutrial_Click()
    FRMTRIAL.Show
    FRMTRIAL.SetFocus
End Sub

Private Sub MNUTRNX_Click()
    FRMTRNXREG.Show
    FRMTRNXREG.SetFocus
End Sub

Private Sub MNUUSER_Click()
    frmUserMast.Show
    frmUserMast.SetFocus
End Sub

Private Sub mnuvehicle_Click()
    FrmVehMaster.Show
    FrmVehMaster.SetFocus
End Sub

Private Sub mnuwsale_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS3.Show
            FRMPOS3.SetFocus
        Else
            FRMSALES2.Show
            FRMSALES2.SetFocus
        End If
    Else
        FRMGST1.Show
        FRMGST1.SetFocus
    End If
End Sub

Private Sub mnuwsale1_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    
    If MDIMAIN.lblgst.Caption = "C" Or MDIMAIN.lblgst.Caption = "N" Then
        If frmLogin.rs!Level = "5" Then
            FRMPOS2.Show
            FRMPOS2.SetFocus
        Else
            FRMSALES1.Show
            FRMSALES1.SetFocus
        End If
    Else
        FRMGST.Show
        FRMGST.SetFocus
    End If
End Sub

Private Sub MNUYEAR_Click()
    If Forms.COUNT > 1 Then
        MsgBox "Please close all the opened windows and try", vbOKOnly, "Financial Year"
        Exit Sub
    End If
    FrmYear.Show
    FrmYear.SetFocus
End Sub

Private Sub mnuzero_Click()
    FrmZeroStk.Show
    FrmZeroStk.SetFocus
End Sub

Private Sub MNYLND_Click()
    FRMLendReg.Show
    FRMLendReg.SetFocus
End Sub

Private Sub MunUPurchase_Click()
    frmLP2.Show
    frmLP2.SetFocus
End Sub

Private Sub PGAYR_Click()
    FRMPaymntreg.Show
    FRMPaymntreg.SetFocus
End Sub

Private Sub PR_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMPURCAHSEREGISTER.Show
    FRMPURCAHSEREGISTER.SetFocus
End Sub

Private Sub RSTBILLS_Click()
    Dim rstTRANX As ADODB.Recordset
    Dim rststock As ADODB.Recordset
    Dim RSTITEM As ADODB.Recordset
    Dim rstCust As ADODB.Recordset
    Dim RSTTOT As ADODB.Recordset
    Dim TOT_AMT As Double
    Dim i As Long
    On Error GoTo ERRHAND
    If MsgBox("THIS MAY TAKE SEVERAL MINUTES TO FINISH DEPENDING ON THE QTY OF ITEMS!!! " & Chr(13) & "DO YOU WANT TO CONTINUE....", vbYesNo) = vbNo Then Exit Sub
    Screen.MousePointer = vbHourglass
    vbalProgressBar1.Visible = True
    vbalProgressBar1.Value = 0
    vbalProgressBar1.ShowText = True
    vbalProgressBar1.Text = "PLEASE WAIT..."
    
    Set rstTRANX = New ADODB.Recordset
    rstTRANX.Open "SELECT DISTINCT VCH_NO, TRX_TYPE, TRX_YEAR FROM TRXFILE where TRX_TYPE <> 'DN' ORDER BY TRX_YEAR DESC, VCH_NO DESC", db, adOpenForwardOnly
    vbalProgressBar1.Min = 0
    If rstTRANX.RecordCount > 0 Then vbalProgressBar1.Max = rstTRANX.RecordCount
    i = 0
    Do Until rstTRANX.EOF
        Set rststock = New ADODB.Recordset
        rststock.Open "SELECT * from TRXMAST where VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
        If (rststock.EOF And rststock.BOF) Then
            Set RSTITEM = New ADODB.Recordset
            RSTITEM.Open "SELECT * FROM TRXFILE where VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenForwardOnly
            If Not (RSTITEM.EOF And RSTITEM.BOF) Then
                i = i + 1
                rststock.AddNew
                rststock!VCH_NO = RSTITEM!VCH_NO
                rststock!TRX_YEAR = RSTITEM!TRX_YEAR
                rststock!TRX_TYPE = RSTITEM!TRX_TYPE
                rststock!VCH_DATE = RSTITEM!VCH_DATE
                rststock!sys_name = system_name
                
                Set rstCust = New ADODB.Recordset
                rstCust.Open "SELECT * FROM CUSTMAST where ACT_NAME = '" & Mid(RSTITEM!VCH_DESC, 15) & "' ", db, adOpenForwardOnly
                If Not (rstCust.EOF And rstCust.BOF) Then
                    If rstCust!ACT_CODE = "130000" Or rstCust!ACT_CODE = "130001" Then
                        rststock!POST_FLAG = "Y"
                    Else
                        rststock!POST_FLAG = "N"
                    End If
                    rststock!TIN = rstCust!KGST
                    rststock!CUST_IGST = IIf(IsNull(rstCust!CUST_IGST), "N", rstCust!CUST_IGST)
                    rststock!ACT_CODE = rstCust!ACT_CODE
                    rststock!ACT_NAME = rstCust!ACT_NAME
                    rststock!phone = rstCust!TELNO
                    rststock!BILL_NAME = rstCust!ACT_NAME
                    rststock!BILL_ADDRESS = rstCust!Address
'                    rststock!BR_CODE = RSTCUST!CUST_CODE
'                    rststock!BR_NAME = RSTCUST!CUST_CODE
'                    rststock!BANK_CODE = BANKCODE
                    rststock!cr_days = rstCust!PYMT_LIMIT
                Else
                    rststock!POST_FLAG = "Y"
                    rststock!TIN = ""
                    rststock!CUST_IGST = "N"
                    rststock!ACT_CODE = "130000"
                    rststock!ACT_NAME = "CASH"
                    rststock!BILL_NAME = "CASH"
                    rststock!BILL_ADDRESS = Mid(RSTITEM!VCH_DESC, 15)
                    rststock!cr_days = 0
                End If
                rstCust.Close
                Set rstCust = Nothing
                rststock!UID_NO = ""
                rststock!DISCOUNT = 0
                rststock!DISC_PERS = 0
                rststock!ADD_AMOUNT = 0
                rststock!ROUNDED_OFF = 0
                rststock!BILL_FLAG = "Y"
                rststock!TERMS = ""
                'VCH_DESC
                
                TOT_AMT = 0
                Set RSTTOT = New ADODB.Recordset
                RSTTOT.Open "SELECT SUM(TRX_TOTAL) FROM TRXFILE WHERE VCH_NO = " & rstTRANX!VCH_NO & " AND TRX_YEAR = '" & rstTRANX!TRX_YEAR & "' AND TRX_TYPE = '" & rstTRANX!TRX_TYPE & "' ", db, adOpenStatic, adLockReadOnly
                If Not (RSTTOT.EOF And RSTTOT.BOF) Then
                    TOT_AMT = IIf(IsNull(RSTTOT.Fields(0)), 0, RSTTOT.Fields(0))
                End If
                RSTTOT.Close
                Set RSTTOT = Nothing
                
                rststock!VCH_AMOUNT = TOT_AMT
                rststock!NET_AMOUNT = TOT_AMT
                rststock!gross_amt = 0
               
                rststock.Update
            End If
            RSTITEM.Close
            Set RSTITEM = Nothing
        End If
        rststock.Close
        Set rststock = Nothing
        
        'db.Execute "Update RTRXFILE SET P_RETAIL = " & rstTRANX!P_RETAIL & " WHERE ITEM_CODE = '" & rstTRANX!ITEM_CODE & "' "
        
        vbalProgressBar1.Value = vbalProgressBar1.Value + 1
        rstTRANX.MoveNext
    Loop
    rstTRANX.Close
    Set rstTRANX = Nothing
    
    vbalProgressBar1.Text = "Successfully Completed"
    Screen.MousePointer = vbNormal
    MsgBox i & " ITEMS", , "EzBiz"
    MsgBox "Completed successfully", vbOKOnly, "EzBiz"
    vbalProgressBar1.Text = ""
    Exit Sub
ERRHAND:
    vbalProgressBar1.Value = 0
    vbalProgressBar1.Text = ""
    Screen.MousePointer = vbNormal
    MsgBox err.Description
End Sub

Private Sub SR_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMBILLPRINT.Show
    FRMBILLPRINT.SetFocus
End Sub

Private Sub StatusBar_PanelDblClick(ByVal Panel As ComctlLib.Panel)
    If frmLogin.rs!Level = "2" Or frmLogin.rs!Level = "5" Then Exit Sub
    If TxtDUP.Visible = True Then
        TxtDUP.Text = ""
        TxtDUP.Visible = False
    Else
        If Panel.index = 5 Then
            TxtDUP.Visible = True
            TxtDUP.SetFocus
        Else
            TxtDUP.Text = ""
            TxtDUP.Visible = False
        End If
    End If
End Sub

Private Sub STKADJUST_KeyDown(KeyCode As MSForms.ReturnInteger, Shift As Integer)
    Dim i As Integer
    i = KeyCode
    If Forms.COUNT = 1 And Shift = vbCtrlMask Then Call Keydown(i)
    
    Exit Sub
End Sub

Private Sub STKADJUST_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub STKMVMNT_Click()
    FrmStkmovmnt.Show
    FrmStkmovmnt.SetFocus
End Sub

Private Sub CmdExit_Click()
'    If IsFormLoaded(frmreport) Then
'        Unload frmreport
'    End If
'    If Forms.COUNT = 1 Then CLOSEALL = 0
    'CLOSEALL = 0
    Unload Me
'    If IsFormLoaded(Me) = True Then
'        Screen.MousePointer = vbNormal
'        MsgBox "Please save and close the opened window", vbOKOnly, "EzBiz"
'    End If
'    db.Close
'    DB.Close
'    Unload Me
End Sub

Private Sub CMDEXIT_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub cmdpurchase_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    If MDIMAIN.LBLSHOPRT.Caption = "Y" Then
        frmLPS.Show
        frmLPS.SetFocus
    Else
        If MDIMAIN.lblcategory.Caption = "Y" Then
            frmLP.Show
            frmLP.SetFocus
        Else
            frmLP1.Show
            frmLP1.SetFocus
        End If
    End If
End Sub

Private Sub cmdpurchase_KeyPress(KeyAscii As MSForms.ReturnInteger)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub PCTMENU_GotFocus()
    On Error Resume Next
    MDIMAIN.CmdRetailBill.SetFocus
End Sub

Private Sub PCTMENU_KeyPress(KeyAscii As Integer)
    Dim i As Integer
    i = KeyAscii
    If Forms.COUNT = 1 Then Call Keypress(i)
End Sub

Private Sub STKADJUST_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    FRMSTKSUMRy.Show
    FRMSTKSUMRy.SetFocus
End Sub

Private Function Keypress(key As Integer)
    
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
    
    Select Case key
            Case vbKeyD, Asc("d")
                cmdduplicate.Tag = key
            Case vbKeyU, Asc("u")
                CMD62.Tag = key
            Case vbKeyP, Asc("p")
'                If exp_flag = True Then
'                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
'                    Call errcodes(Val(lblec.Caption))
'                    Exit Function
'                End If
'                If (Val(cmdduplicate.Tag) = 68 Or Val(cmdduplicate.Tag) = 100) And (Val(CMD62.Tag) = 85 Or Val(CMD62.Tag) = 117) Then
''                    pctmenu2.Visible = True
''                    PCTMENU.Visible = False
'                    'Me.Caption = "Ez..Biz INVENTORY - ESTIMATE"
'                    'Me.Picture = Rptpath & "Geo.jpg" '
'                    CMDDUPPURCHASE_Click
'                End If
                cmdduplicate.Tag = ""
            Case vbKeyS, Asc("s")
                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                If (Val(cmdduplicate.Tag) = 68 Or Val(cmdduplicate.Tag) = 100) And (Val(CMD62.Tag) = 85 Or Val(CMD62.Tag) = 117) Then
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing
                    cmdduplicate_Click
                End If
                cmdduplicate.Tag = ""
            
'            Case vbKeyB, Asc("b")
'                cmdduplicate.Tag = Key
'            Case vbKeyK, Asc("k")
'                If Val(cmdduplicate.Tag) = 66 Or Val(cmdduplicate.Tag) = 98 Then
'                    Command1_Click
'                End If
'                cmdduplicate.Tag = ""
            Case vbKeyE, Asc("e")
                cmdduplicate.Tag = key
            Case vbKeyZ, Asc("z")
                CMD62.Tag = key
            Case vbKeyB, Asc("b")
                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                If (Val(cmdduplicate.Tag) = 101 Or Val(cmdduplicate.Tag) = 69) And (Val(CMD62.Tag) = 122 Or Val(CMD62.Tag) = 90) Then
                    frmeasybill.Show
                    frmeasybill.SetFocus
                End If
                cmdduplicate.Tag = ""
            Case vbKeyR, Asc("r")
                cmdduplicate.Tag = key
            Case vbKeyU, Asc("u")
                CMD62.Tag = key
            Case vbKeyK, Asc("k")
                CMD62.Tag = key
'            Case vbKeyE, Asc("e")
'                CMD62.Tag = Key
            Case vbKeyY, Asc("y")
                If (Val(CMD62.Tag) = 107 Or Val(CMD62.Tag) = 75) And (Val(cmdduplicate.Tag) = 101 Or Val(cmdduplicate.Tag) = 69) Then
                    If (frmLogin.rs!Level = "0" Or frmLogin.rs!Level = "4") Then
                        Dim ACT_KEY1, ACT_KEY2 As String
                        Dim MD5 As New clsMD5
                        
                        ACT_KEY1 = Val(GetUniqueCode) * 555
                        
                        ACT_KEY2 = UCase(MD5.DigestStrToHexStr(ACT_KEY1))
                        ACT_KEY2 = ACT_KEY2 & UCase(MD5.DigestStrToHexStr(ACT_KEY2))
                        ACT_KEY2 = Mid(ACT_KEY2, 24, 10) & Mid(ACT_KEY2, 1, 5)
                        MDIMAIN.Enabled = False
                        Dim sql As String
                        Dim TRXFILE As ADODB.Recordset
                        Set TRXFILE = New ADODB.Recordset
                        sql = "select * from act_ky WHERE ACT_CODE= '" & ACT_KEY2 & "'"
                        TRXFILE.Open sql, db, adOpenKeyset, adLockPessimistic
                        If Not (TRXFILE.BOF And TRXFILE.EOF) Then
                            FrmKey.Show
                            FrmKey.SetFocus
                            FrmKey.lblINSID.Caption = ACT_KEY1
                            TRXFILE.Close
                            Set TRXFILE = Nothing
                            Exit Function
                        Else
                            FrmKey.Show
                            FrmKey.SetFocus
                        End If
                        TRXFILE.Close
                        Set TRXFILE = Nothing
                    End If
                End If
            Case vbKeyN, Asc("n")
                If (Val(cmdduplicate.Tag) = 114 Or Val(cmdduplicate.Tag) = 82) And (Val(CMD62.Tag) = 117 Or Val(CMD62.Tag) = 85) Then
                    Frmexpkey.Show
                    Frmexpkey.SetFocus
                    Exit Function
                End If
            Case Else
                cmdduplicate.Tag = ""
                'cmdEstimate.Tag = ""
    End Select
    
    Exit Function
ERRHAND:
    MsgBox err.Description
End Function

Private Sub MNUACCOUNTS_Click()
    If exp_flag = True Then
        'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
        Call errcodes(Val(lblec.Caption))
        Exit Sub
    End If
    If (frmLogin.rs!Level <> "0" And frmLogin.rs!Level <> "4") Then Exit Sub
    On Error Resume Next
    frmDaybook.Show
    frmDaybook.SetFocus
End Sub

Private Sub cmdtest_Click()
'    Dim RSTTRXFILE As ADODB.Recordset
'    Dim rststock As ADODB.Recordset
'    Dim rstitem3 As ADODB.Recordset
'    Dim rstlink1 As ADODB.Recordset
'    Dim rstlink2, RSTRTRXFILE As ADODB.Recordset
'    Dim i As Double
'
'
'    Exit Sub
'
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockReadOnly, adCmdText
'    Do Until rststock.EOF
'        i = rststock!CLOSE_QTY
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ", db, adOpenStatic, adLockOptimistic, adCmdText
'        Do Until RSTRTRXFILE.EOF
'            RSTRTRXFILE!BAL_QTY = 0
'            RSTRTRXFILE.Update
'            RSTRTRXFILE.MoveNext
'        Loop
'        RSTRTRXFILE.Close
'        Set RSTRTRXFILE = Nothing
'
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY VCH_DATE", db, adOpenStatic, adLockOptimistic, adCmdText
'        If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
'            RSTRTRXFILE.MoveLast
'            RSTRTRXFILE!BAL_QTY = i
'            RSTRTRXFILE.Update
'        End If
'        RSTRTRXFILE.Close
'        Set RSTRTRXFILE = Nothing
'
'        i = 0
'
'        rststock.MoveNext
'    Loop
'
'
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'   Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM TRXMAST", db, adOpenStatic, adLockReadOnly, adCmdText
'    Do Until rststock.EOF
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT * FROM TRXFILE WHERE TRX_TYPE = '" & rststock!TRX_TYPE & "' AND VCH_NO = " & rststock!VCH_NO & " ORDER BY LINE_NO", db, adOpenStatic, adLockOptimistic, adCmdText
'        'If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
'        Do Until RSTRTRXFILE.EOF
'            RSTRTRXFILE!M_USER_ID = rststock!ACT_CODE
'            RSTRTRXFILE.Update
'            RSTRTRXFILE.MoveNext
'        Loop
'        'End If
'        'i = i + 1
'        rststock.MoveNext
'    Loop
'    RSTRTRXFILE.Close
'    Set RSTRTRXFILE = Nothing
'
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until rststock.EOF
'        rststock!ITEM_NAME = rststock!Category & " " & rststock!ITEM_NAME
'        rststock.Update
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'    i = 1
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until rststock.EOF
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
'        If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
'        'Do Until RSTRTRXFILE.EOF
'            rststock!UNIT = RSTRTRXFILE!LINE_DISC
'            rststock!ITEM_COST = RSTRTRXFILE!ITEM_COST
'            rststock!P_RETAIL = RSTRTRXFILE!P_RETAIL
'            rststock!P_WS = RSTRTRXFILE!P_WS
'            rststock!P_CRTN = RSTRTRXFILE!P_CRTN
'            rststock!P_RETAIL = RSTRTRXFILE!P_VAN
'            rststock!MRP = RSTRTRXFILE!MRP
'            'rststock!MRP_BT = IIf(IsNull(RSTRTRXFILE!MRP_BT), 0, RSTRTRXFILE!MRP_BT)
'            rststock!SALES_TAX = RSTRTRXFILE!SALES_TAX
'            rststock!PTR = RSTRTRXFILE!PTR
'            rststock!COM_PER = RSTRTRXFILE!COM_PER
'            rststock!COM_AMT = RSTRTRXFILE!COM_PER
'            rststock!COM_FLAG = RSTRTRXFILE!COM_FLAG
'            rststock!CRTN_PACK = RSTRTRXFILE!CRTN_PACK
'            rststock.Update
'            'RSTRTRXFILE.MoveNext
'        'Loop
'        End If
'        'i = i + 1
'        rststock.MoveNext
'    Loop
'    RSTRTRXFILE.Close
'    Set RSTRTRXFILE = Nothing
'
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'    i = 1
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until rststock.EOF
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT * FROM RTRXFILE WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "' AND TRX_TYPE='PI' ORDER BY VCH_DATE DESC", db, adOpenStatic, adLockReadOnly
'        If Not (RSTRTRXFILE.EOF And RSTRTRXFILE.BOF) Then
'        'Do Until RSTRTRXFILE.EOF
'            rststock!UNIT = RSTRTRXFILE!LINE_DISC
'            rststock!ITEM_COST = RSTRTRXFILE!ITEM_COST
'            rststock!MRP = RSTRTRXFILE!MRP
'            rststock!MRP_BT = IIf(IsNull(RSTRTRXFILE!MRP_BT), 0, RSTRTRXFILE!MRP_BT)
'            rststock!SALES_TAX = RSTRTRXFILE!SALES_TAX
'            rststock!PTR = RSTRTRXFILE!PTR
'
'
'
'            rststock.Update
'            'RSTRTRXFILE.MoveNext
'        'Loop
'        End If
'        'i = i + 1
'        rststock.MoveNext
'    Loop
'    RSTRTRXFILE.Close
'    Set RSTRTRXFILE = Nothing
'
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'    Set rststock = New ADODB.Recordset
'    rststock.Open "SELECT * From RTRXFILE ORDER BY ITEM_CODE", db, adOpenStatic, adLockReadOnly
'    Do Until rststock.EOF
'        Set RSTRTRXFILE = New ADODB.Recordset
'        RSTRTRXFILE.Open "SELECT *  FROM ITEMMAST WHERE ITEM_CODE = '" & rststock!ITEM_CODE & "'", db, adOpenStatic, adLockOptimistic, adCmdText
'        With RSTRTRXFILE
'            If Not (.EOF And .BOF) Then
'                !P_RETAIL = rststock!P_RETAIL
'                !P_VAN = rststock!P_VAN
'                !P_CRTN = rststock!P_CRTN
'                !P_WS = rststock!P_WS
'                RSTRTRXFILE.Update
'            End If
'        End With
'        RSTRTRXFILE.Close
'        Set RSTRTRXFILE = Nothing
'
'        rststock.MoveNext
'    Loop
'    rststock.Close
'    Set rststock = Nothing
'
'    Exit Sub
'    db.Execute "delete FROM ATRXFILE"
'    db.Execute "delete FROM ATRXSUB"
'    db.Execute "delete FROM BANKCODE"
'    db.Execute "delete FROM BANKLETTERS"
'    db.Execute "delete FROM BONUSMAST"
'    db.Execute "delete FROM CANCINV"
'    db.Execute "delete FROM CHQMAST"
'    db.Execute "delete FROM DAMAGED"
'    db.Execute "delete FROM FQTYLIST"
'    db.Execute "delete FROM ORDISSUE"
'    db.Execute "delete FROM ORDMAST"
'    db.Execute "delete FROM ORDSUB"
'    db.Execute "delete FROM PASSWORDS"
'    db.Execute "delete FROM POMAST"
'    db.Execute "delete FROM POSUB"
'    db.Execute "delete FROM PRICETABLE"
'    db.Execute "delete FROM QTNMAST"
'    db.Execute "delete FROM QTNSUB"
'    db.Execute "delete FROM REORDER"
'    db.Execute "delete FROM RTRXFILE"
'    db.Execute "delete FROM ITEMMAST"
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * From ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until RSTTRXFILE.EOF
'        RSTTRXFILE!BIN_LOCATION = Mid(RSTTRXFILE!ITEM_NAME, 1, 1)
'        RSTTRXFILE!SALES_TAX = 0
'        RSTTRXFILE!CST = 0
'        RSTTRXFILE!OPEN_QTY = 0
'        RSTTRXFILE!OPEN_VAL = 0
'        RSTTRXFILE!RCPT_QTY = 0
'        RSTTRXFILE!RCPT_VAL = 0
'        RSTTRXFILE!ISSUE_QTY = 0
'        RSTTRXFILE!ISSUE_VAL = 0
'        RSTTRXFILE!CLOSE_QTY = 0
'        RSTTRXFILE!CLOSE_VAL = 0
'        RSTTRXFILE!DAM_QTY = 0
'        RSTTRXFILE!DAM_VAL = 0
'        RSTTRXFILE!DISC = 0
'        RSTTRXFILE.Update
'        RSTTRXFILE.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
'
'    'db.Execute "delete FROM RTRXFILE WHERE VCH_DATE <# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND BAL_QTY <=0"
'    db.Execute "delete FROM REPLCN"
'    db.Execute "delete FROM TEMPCN"
'    db.Execute "delete FROM TRANSMAST"
'    db.Execute "delete FROM TRANSSUB"
'    db.Execute "delete FROM TRXEXPENSE"
'    db.Execute "delete FROM TRXFILE"
'    db.Execute "delete FROM TRXMAST"
'    db.Execute "delete FROM TRXSUB"
'    db.Execute "delete FROM VANSTOCK"
'
'    db.Execute "delete FROM ATRXFILE"
'    db.Execute "delete FROM CRDTPYMT"
'    db.Execute "delete FROM RTRXFILE"
'    db.Execute "delete FROM SALERETURN"
'    db.Execute "delete FROM TRANSMASTWO"
'    db.Execute "delete FROM TRXEXPENSE"
'    db.Execute "delete FROM TRXFILE"
'    db.Execute "delete FROM TRXMAST"
'    db.Execute "delete FROM TRXSUB"
'    db.Execute "delete FROM TRXWOBILL"
'
'
'
'
'
'    db.Execute "delete FROM BAKEXPIRY"
'    db.Execute "delete FROM BILLDETAILS"
'    db.Execute "delete FROM CRDTPYMT"
'    db.Execute "delete FROM CRDTPYMT1"
'    db.Execute "delete FROM DBTPYMT"
'    db.Execute "delete FROM Delivery"
'    db.Execute "delete FROM DUMMYBILL"
'    db.Execute "delete FROM EXPIRY"
'    db.Execute "delete FROM EXPLIST"
'    db.Execute "delete FROM EXPSORT"
'    db.Execute "delete FROM MinStock"
'    db.Execute "delete FROM NONRCVD"
'    db.Execute "delete FROM P_Rate"
'    db.Execute "delete FROM PURCAHSERETURN"
'    db.Execute "delete FROM RTRXFILE"
'    db.Execute "delete FROM SALEBILL"
'    db.Execute "delete FROM SALERETURN"
'    db.Execute "delete FROM SALESLEDGER"
'    db.Execute "delete FROM SALESREG"
'    db.Execute "delete FROM SALESREG2"
'    db.Execute "delete FROM SelDist"
'    db.Execute "delete FROM SLIP_REG"
'    db.Execute "delete FROM Stock"
'    db.Execute "delete FROM STOCKLESS"
'    db.Execute "delete FROM TEMPCN"
'    db.Execute "delete FROM TEMPTRX"
'    db.Execute "delete FROM TRNXRCPT"
'    db.Execute "delete FROM TmpOrderlist"
'    db.Execute "delete FROM TRXFILE"
'    db.Execute "delete FROM TRXMAST"
'    db.Execute "delete FROM TRXWOBILL"
'
'
    
    
'''    Set RSTTRXFILE = New ADODB.Recordset
'''    RSTTRXFILE.Open "Select * From TRXMAST order by VCH_NO", DB, adOpenStatic, adLockOptimistic, adCmdText
'''    Do Until RSTTRXFILE.EOF
'''        Set rstitem2 = New ADODB.Recordset
'''        rstitem2.Open "Select * From TRXMAST WHERE VCH_NO = " & RSTTRXFILE!T_VCH_NO & "", db, adOpenStatic, adLockReadOnly
'''        Do Until rstitem2.EOF
'''            RSTTRXFILE!VCH_DATE = RSTTRXFILE!VCH_DATE
'''            RSTTRXFILE.Update
'''            rstitem2.MoveNext
'''        Loop
'''        RSTTRXFILE.MoveNext
'''    Loop
'''    RSTTRXFILE.Close
'''    Set RSTTRXFILE = Nothing
    

''
''
'''    DTFROM.Value = "01/01/2012"
'''    db.Execute "delete FROM RTRXFILE WHERE VCH_DATE <# " & Format(DTFROM.Value, "MM,DD,YYYY") & " # AND BAL_QTY <=0"
'''
'''    Set RSTTRXFILE = New ADODB.Recordset
'''    RSTTRXFILE.Open "Select VCH_AMOUNT, VCH_NO, DISCOUNT From TRANSMAST WHERE TRX_TYPE='PI' ORDER BY VCH_NO", db, adOpenStatic, adLockOptimistic, adCmdText
'''    Do Until RSTTRXFILE.EOF
'''        i = 0
'''        Set rstitem3 = New ADODB.Recordset
'''        rstitem3.Open "Select * From RTRXFILE WHERE TRX_TYPE='PI' AND VCH_NO = " & RSTTRXFILE!VCH_NO & "", db, adOpenForwardOnly
'''        Do Until rstitem3.EOF
'''            i = i + rstitem3!TRX_TOTAL
'''            rstitem3.MoveNext
'''        Loop
'''        rstitem3.Close
'''        Set rstitem3 = Nothing
'''
'''        RSTTRXFILE!VCH_AMOUNT = Round(i - RSTTRXFILE!DISCOUNT)
'''        RSTTRXFILE.Update
'''        RSTTRXFILE.MoveNext
'''    Loop
'''    RSTTRXFILE.Close
'''    Set RSTTRXFILE = Nothing
''
''
''
'''    On Error GoTo errHand
'''    Screen.MousePointer = vbHourglass
'''    Set rstitem1 = New ADODB.Recordset
'''    rstitem1.Open "SELECT DISTINCT ITEM_CODE from RTRXFILE", db, adOpenForwardOnly
'''    Do Until rstitem1.EOF
'''        Set rstitem2 = New ADODB.Recordset
'''        rstitem2.Open "SELECT * from ITEMMAST WHERE ITEM_CODE = '" & rstitem1!ITEM_CODE & "'", db, adOpenForwardOnly
'''        If Not (rstitem2.EOF And rstitem2.BOF) Then
'''            Set rstitem3 = New ADODB.Recordset
'''            rstitem3.Open "SELECT * from ITEMMAST", DB, adOpenStatic, adLockOptimistic, adCmdText
'''            rstitem3.AddNew
'''            rstitem3!ITEM_CODE = rstitem2!ITEM_CODE
'''            rstitem3!ITEM_NAME = rstitem2!ITEM_NAME
'''            rstitem3!UNIT = rstitem2!UNIT
'''            rstitem3!CATEGORY = rstitem2!CATEGORY
'''            rstitem3!ITEM_COST = rstitem2!ITEM_COST
'''            rstitem3!MRP = rstitem2!MRP
'''            rstitem3!SALES_TAX = rstitem2!SALES_TAX
'''            rstitem3!PTR = rstitem2!PTR
'''            rstitem3!BIN_LOCATION = rstitem2!BIN_LOCATION
'''            rstitem3!CST = rstitem2!CST
'''            rstitem3!OPEN_QTY = rstitem2!OPEN_QTY
'''            rstitem3!OPEN_VAL = rstitem2!OPEN_VAL
'''            rstitem3!RCPT_QTY = rstitem2!RCPT_QTY
'''            rstitem3!RCPT_VAL = rstitem2!RCPT_VAL
'''            rstitem3!ISSUE_QTY = rstitem2!ISSUE_QTY
'''            rstitem3!ISSUE_VAL = rstitem2!ISSUE_VAL
'''            rstitem3!CLOSE_QTY = rstitem2!CLOSE_QTY
'''            rstitem3!CLOSE_VAL = rstitem2!CLOSE_VAL
'''            rstitem3!DAM_QTY = rstitem2!DAM_QTY
'''            rstitem3!DAM_VAL = rstitem2!DAM_VAL
'''            rstitem3!SCHEDULE = rstitem2!SCHEDULE
'''            rstitem3!DISC = rstitem2!DISC
'''            rstitem3!REORDER_QTY = rstitem2!REORDER_QTY
'''            rstitem3!MANUFACTURER = rstitem2!MANUFACTURER
'''            rstitem3!SUPPLIER = rstitem2!SUPPLIER
'''            rstitem3!Remarks = rstitem2!Remarks
'''            rstitem3!CREATE_DATE = rstitem2!CREATE_DATE
'''            rstitem3!C_USER_ID = rstitem2!C_USER_ID
'''            rstitem3!MODIFY_DATE = rstitem2!MODIFY_DATE
'''            rstitem3!M_USER_ID = rstitem2!M_USER_ID
'''
'''            rstitem3.Update
'''            rstitem3.Close
'''            Set rstitem3 = Nothing
'''
'''            Set rstlink1 = New ADODB.Recordset
'''            rstlink1.Open "SELECT * from PRODLINK WHERE ITEM_CODE = '" & rstitem1!ITEM_CODE & "'", db, adOpenForwardOnly
'''            Set rstlink2 = New ADODB.Recordset
'''            rstlink2.Open "SELECT * from PRODLINK", DB, adOpenStatic, adLockOptimistic, adCmdText
'''            Do Until rstlink1.EOF
'''                rstlink2.AddNew
'''                rstlink2!ITEM_CODE = rstlink1!ITEM_CODE
'''                rstlink2!ITEM_NAME = rstlink1!ITEM_NAME
'''                rstlink2!RQTY = rstlink1!RQTY
'''                rstlink2!ITEM_COST = rstlink1!ITEM_COST
'''                rstlink2!MRP = rstlink1!MRP
'''                rstlink2!PTR = rstlink1!PTR
'''                rstlink2!SALES_PRICE = rstlink1!SALES_PRICE
'''                rstlink2!SALES_TAX = rstlink1!SALES_TAX
'''                rstlink2!UNIT = rstlink1!UNIT
'''                rstlink2!Remarks = rstlink1!Remarks
'''                rstlink2!ORD_QTY = rstlink1!ORD_QTY
'''                rstlink2!CST = rstlink1!CST
'''                rstlink2!ACT_CODE = rstlink1!ACT_CODE
'''                rstlink2!CREATE_DATE = rstlink1!CREATE_DATE
'''                rstlink2!C_USER_ID = rstlink1!C_USER_ID
'''                rstlink2!MODIFY_DATE = rstlink1!MODIFY_DATE
'''                rstlink2!M_USER_ID = rstlink1!M_USER_ID
'''                rstlink2!CHECK_FLAG = rstlink1!CHECK_FLAG
'''                rstlink2!SITEM_CODE = rstlink1!SITEM_CODE
'''
'''                rstlink2.Update
'''                rstlink1.MoveNext
'''            Loop
'''
'''            rstlink2.Close
'''            Set rstlink2 = Nothing
'''            rstlink1.Close
'''            Set rstlink1 = Nothing
'''        End If
'''        rstitem2.Close
'''        Set rstitem2 = Nothing
'''
'''        rstitem1.MoveNext
'''    Loop
'''    rstitem1.Close
'''    Set rstitem1 = Nothing
'''    Screen.MousePointer = vbNormal
'''    Exit Sub
'''errHand:
'''    Screen.MousePointer = vbNormal
'''    MsgBox Err.Description
'''End Sub
'''
'''''''Private Sub Command1_Click()
'''''''    Dim RSTFILE As ADODB.Recordset
'''''''    Dim RSTFILE2 As ADODB.Recordset
'''''''    Dim RSTFILE3 As ADODB.Recordset
'''''''    Dim i As Long
'''''''    Dim n As Long
'''''''
'''''''''' '''   //////SUPPLIER LEDGER//////
''''''''''    Set RSTFILE = New ADODB.Recordset
''''''''''    RSTFILE.Open "Select * From Phm003 ", Conn2, adOpenForwardOnly
''''''''''    Do Until RSTFILE.EOF
''''''''''        Set RSTFILE2 = New ADODB.Recordset
''''''''''        RSTFILE2.Open "Select * From CUSTMAST", db, adOpenStatic, adLockOptimistic, adCmdText
''''''''''        RSTFILE2.AddNew
''''''''''        RSTFILE2!ACT_CODE = "311" & RSTFILE!SUPCD
''''''''''        RSTFILE2!ACT_NAME = RSTFILE!SUPNM
''''''''''        RSTFILE2!ADDRESS = RSTFILE!SADD1 & ", " & RSTFILE!SADD2
''''''''''        RSTFILE2!TELNO = Mid(RSTFILE!PHONE, 1, 10)
''''''''''        RSTFILE2!KGST = RSTFILE!KGSTNO
''''''''''        RSTFILE2!CST = ""
''''''''''        RSTFILE2!DL_NO = ""
''''''''''        RSTFILE2!OPEN_DB = 0
''''''''''        RSTFILE2!OPEN_CR = 0
''''''''''        RSTFILE2!YTD_DB = 0
''''''''''        RSTFILE2!YTD_CR = 0
''''''''''        RSTFILE2!CUST_TYPE = ""
''''''''''        RSTFILE2!C_USER_ID = "SM"
''''''''''
''''''''''        RSTFILE2.Update
''''''''''        RSTFILE2.Close
''''''''''        Set RSTFILE2 = Nothing
''''''''''        RSTFILE.MoveNext
''''''''''    Loop
''''''''''    RSTFILE.Close
''''''''''    Set RSTFILE = Nothing
''''''''''
''''''''''  '  ///////////ITEMMASTER/////////
''''''''''    Set RSTFILE = New ADODB.Recordset
''''''''''    RSTFILE.Open "Select * From Phm001", Conn2, adOpenForwardOnly
''''''''''    Do Until RSTFILE.EOF
''''''''''        Set RSTFILE2 = New ADODB.Recordset
''''''''''        RSTFILE2.Open "Select * From ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
''''''''''        RSTFILE2.AddNew
''''''''''        RSTFILE2!ITEM_CODE = RSTFILE!ITCD
''''''''''        RSTFILE2!ITEM_NAME = RSTFILE!ITNM
''''''''''        RSTFILE2!UNIT = RSTFILE!PACKQTY
''''''''''        RSTFILE2!CATEGORY = "GENERAL"
''''''''''        RSTFILE2!ITEM_COST = RSTFILE!PACKRATE
''''''''''        RSTFILE2!MRP = RSTFILE!PACKMRP / RSTFILE!PACKQTY
''''''''''        RSTFILE2!SALES_TAX = RSTFILE!TAXCD
''''''''''        RSTFILE2!PTR = RSTFILE!PACKRATE
''''''''''        RSTFILE2!BIN_LOCATION = Mid(RSTFILE!ITNM, 1, 1)
''''''''''        RSTFILE2!CST = RSTFILE!CST
''''''''''        RSTFILE2!OPEN_QTY = 0
''''''''''        RSTFILE2!OPEN_VAL = 0
''''''''''        RSTFILE2!RCPT_QTY = 0
''''''''''        RSTFILE2!RCPT_VAL = 0
''''''''''        RSTFILE2!ISSUE_QTY = 0
''''''''''        RSTFILE2!ISSUE_VAL = 0
''''''''''        RSTFILE2!CLOSE_QTY = RSTFILE!BALQTY
''''''''''        RSTFILE2!CLOSE_VAL = 0
''''''''''        RSTFILE2!DAM_QTY = 0
''''''''''        RSTFILE2!DAM_VAL = 0
''''''''''        RSTFILE2!SCHEDULE = "H"
''''''''''        RSTFILE2!DISC = 0
''''''''''        RSTFILE2!REORDER_QTY = RSTFILE!PACKQTY
''''''''''        Set RSTFILE3 = New ADODB.Recordset
''''''''''        RSTFILE3.Open "SELECT * from Phm006 WHERE COCODE = '" & RSTFILE!COCODE & "'", Conn2, adOpenForwardOnly
''''''''''        If Not (RSTFILE3.EOF And RSTFILE3.BOF) Then
''''''''''            RSTFILE2!MANUFACTURER = Mid(RSTFILE3!CONAME, 1, 20)
''''''''''        End If
''''''''''        RSTFILE3.Close
''''''''''        Set RSTFILE3 = Nothing
''''''''''        RSTFILE2!SUPPLIER = ""
''''''''        RSTFILE2!Remarks = RSTFILE!PACKRATE
''''''''        RSTFILE2!CREATE_DATE = Format(Date, "DD/MM/YYYY")
''''''''        RSTFILE2!C_USER_ID = ""
''''''''        RSTFILE2!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
''''''''        RSTFILE2!M_USER_ID = ""
''''''''
''''''''        RSTFILE2.Update
''''''''        RSTFILE2.Close
''''''''        Set RSTFILE2 = Nothing
''''''''        RSTFILE.MoveNext
''''''''    Loop
''''''''    RSTFILE.Close
''''''''    Set RSTFILE = Nothing
'''''
''''''/////////RTRXFILE///////ITEMS WITH BATCH
'''''    i = 0
'''''    Set RSTFILE = New ADODB.Recordset
'''''    RSTFILE.Open "Select * From Phm002 WHERE QTY <>0", Conn2, adOpenForwardOnly
'''''    Do Until RSTFILE.EOF
'''''        Set RSTFILE2 = New ADODB.Recordset
'''''        RSTFILE2.Open "Select * From RTRXFILE", db, adOpenStatic, adLockOptimistic, adCmdText
'''''        RSTFILE2.AddNew
'''''        i = i + 1
'''''        RSTFILE2!TRX_TYPE = "PI"
'''''        RSTFILE2!VCH_NO = i
'''''        Set RSTFILE3 = New ADODB.Recordset
'''''        RSTFILE3.Open "SELECT * from Pht001 WHERE INVNO = '" & RSTFILE!INVNO & "'", Conn2, adOpenForwardOnly
'''''        If Not (RSTFILE3.EOF And RSTFILE3.BOF) Then
'''''            RSTFILE2!VCH_DATE = Format(RSTFILE3!INVDT, "DD/MM/YYYY")
'''''        Else
'''''            RSTFILE2!VCH_DATE = Format("01/01/2001")
'''''        End If
'''''        RSTFILE3.Close
'''''        Set RSTFILE3 = Nothing
'''''
'''''        RSTFILE2!LINE_NO = 1
'''''        RSTFILE2!CATEGORY = "GENERAL"
'''''        RSTFILE2!ITEM_CODE = RSTFILE!ITCD
'''''        RSTFILE2!QTY = RSTFILE!QTY
'''''        RSTFILE2!ITEM_COST = RSTFILE!CPRICE
'''''        RSTFILE2!PTR = RSTFILE!CPRICE
'''''        RSTFILE2!SALES_PRICE = RSTFILE!Rate
'''''        RSTFILE2!SALES_TAX = RSTFILE!TAX
'''''        Set RSTFILE3 = New ADODB.Recordset
'''''        RSTFILE3.Open "SELECT * from Phm001 WHERE ITCD = '" & RSTFILE!ITCD & "'", Conn2, adOpenForwardOnly
'''''        If Not (RSTFILE3.EOF And RSTFILE3.BOF) Then
'''''            RSTFILE2!UNIT = RSTFILE3!PACKQTY
'''''            RSTFILE2!MRP = RSTFILE!Rate * RSTFILE3!PACKQTY
'''''            RSTFILE2!ITEM_NAME = RSTFILE3!ITNM
'''''        End If
'''''        RSTFILE3.Close
'''''        Set RSTFILE3 = Nothing
'''''
'''''        Command1.Tag = "311" & RSTFILE!SUPCD
'''''        Set RSTFILE3 = New ADODB.Recordset
'''''        RSTFILE3.Open "SELECT * from CUSTMAST WHERE ACT_CODE = '" & Command1.Tag & "'", db, adOpenForwardOnly
'''''        If Not (RSTFILE3.EOF And RSTFILE3.BOF) Then
'''''            RSTFILE2!VCH_DESC = "Received From " & RSTFILE3!ACT_NAME
'''''        End If
'''''        RSTFILE3.Close
'''''        Set RSTFILE3 = Nothing
'''''
'''''        RSTFILE2!REF_NO = RSTFILE!BTNO
'''''        RSTFILE2!ISSUE_QTY = 0
'''''        RSTFILE2!CST = 0
'''''        RSTFILE2!BAL_QTY = RSTFILE!QTY
'''''        RSTFILE2!TRX_TOTAL = 0
'''''        RSTFILE2!LINE_DISC = Null
'''''        RSTFILE2!SCHEME = Null
'''''        RSTFILE2!EXP_DATE = RSTFILE!EXPDT
'''''        RSTFILE2!FREE_QTY = 0
'''''        RSTFILE2!CREATE_DATE = Format(Date, "DD/MM/YYYY")
'''''        RSTFILE2!C_USER_ID = ""
'''''        RSTFILE2!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
'''''        RSTFILE2!M_USER_ID = "311" & RSTFILE!SUPCD
'''''        RSTFILE2!CHECK_FLAG = ""
'''''        RSTFILE2!PINV = RSTFILE!INVNO
'''''
'''''        RSTFILE2.Update
'''''        RSTFILE2.Close
'''''        Set RSTFILE2 = Nothing
'''''        RSTFILE.MoveNext
'''''    Loop
'''''    RSTFILE.Close
'''''    Set RSTFILE = Nothing
'''''
'''''''''''/////////RTRXFILE///////STOCK CORRECTION
'''''''    Set RSTITEMMAST = New ADODB.Recordset
'''''''    RSTITEMMAST.Open "SELECT *  FROM ITEMMAST", db, adOpenStatic, adLockOptimistic, adCmdText
'''''''    Do Until RSTITEMMAST.EOF
'''''''        i = 0
'''''''        n = 0
'''''''        Set rststock = New ADODB.Recordset
'''''''        rststock.Open "SELECT BAL_QTY, MRP from RTRXFILE WHERE RTRXFILE.ITEM_CODE = '" & RSTITEMMAST!ITEM_CODE & "'  AND RTRXFILE.BAL_QTY > 0 ORDER BY VCH_NO", db, adOpenForwardOnly
'''''''        Do Until rststock.EOF
'''''''            i = i + rststock!BAL_QTY
'''''''            n = rststock!MRP
'''''''            rststock.MoveNext
'''''''        Loop
'''''''        rststock.Close
'''''''        Set rststock = Nothing
'''''''        RSTITEMMAST!MRP = n
'''''''        RSTITEMMAST!OPEN_QTY = i
'''''''        RSTITEMMAST!OPEN_VAL = 0
'''''''        RSTITEMMAST!RCPT_QTY = 0
'''''''        RSTITEMMAST!RCPT_VAL = 0
'''''''        RSTITEMMAST!ISSUE_QTY = 0
'''''''        RSTITEMMAST!ISSUE_VAL = 0
'''''''        RSTITEMMAST!CLOSE_QTY = i
'''''''        RSTITEMMAST!CLOSE_VAL = 0
'''''''        RSTITEMMAST!DAM_QTY = 0
'''''''        RSTITEMMAST!DAM_VAL = 0
'''''''        RSTITEMMAST.Update
'''''''        RSTITEMMAST.MoveNext
'''''''    Loop
'''''''    RSTITEMMAST.Close
'''''''    Set RSTITEMMAST = Nothing
'''''
'''''     ''''''   /////////PRODUCT LINK///////////
''''''    Set RSTFILE = New ADODB.Recordset
''''''    RSTFILE.Open "Select * From Phm001", Conn2, adOpenForwardOnly
''''''    Do Until RSTFILE.EOF
''''''        Set RSTFILE3 = New ADODB.Recordset
''''''        RSTFILE3.Open "SELECT * from Phm011 WHERE COCODE = '" & RSTFILE!COCODE & "'", Conn2, adOpenForwardOnly
''''''        Do Until RSTFILE3.EOF
''''''            Set RSTFILE2 = New ADODB.Recordset
''''''            RSTFILE2.Open "Select * From PRODLINK", db, adOpenStatic, adLockOptimistic, adCmdText
''''''            RSTFILE2.AddNew
''''''            RSTFILE2!ACT_CODE = "311" & RSTFILE3!SUPCD
''''''
''''''            RSTFILE2!ITEM_CODE = RSTFILE!ITCD
''''''            RSTFILE2!ITEM_NAME = RSTFILE!ITNM
''''''            RSTFILE2!RQTY = 0
''''''            RSTFILE2!ITEM_COST = RSTFILE!PACKRATE
''''''            RSTFILE2!MRP = RSTFILE!PACKMRP
''''''            RSTFILE2!SALES_TAX = RSTFILE!TAXCD
''''''            RSTFILE2!PTR = RSTFILE!PACKRATE
''''''            RSTFILE2!SALES_PRICE = RSTFILE!PACKMRP
''''''            RSTFILE2!UNIT = RSTFILE!PACKQTY
''''''            RSTFILE2!Remarks = RSTFILE!PACKQTY
''''''            RSTFILE2!ORD_QTY = 0
''''''            RSTFILE2!CST = 0
''''''            RSTFILE2!CREATE_DATE = Format(Date, "DD/MM/YYYY")
''''''            RSTFILE2!C_USER_ID = ""
''''''            RSTFILE2!MODIFY_DATE = Format(Date, "DD/MM/YYYY")
''''''            RSTFILE2!M_USER_ID = ""
''''''            RSTFILE2!CHECK_FLAG = "Y"
''''''            RSTFILE2!SITEM_CODE = ""
''''''
''''''            RSTFILE2.Update
''''''            RSTFILE2.Close
''''''            Set RSTFILE2 = Nothing
''''''            RSTFILE3.MoveNext
''''''        Loop
''''''        RSTFILE3.Close
''''''        Set RSTFILE3 = Nothing
''''''
''''''        RSTFILE.MoveNext
''''''    Loop
''''''    RSTFILE.Close
''''''    Set RSTFILE = Nothing
'''''
'''''''''''/////////RTRXFILE///////Payment Details
'''''    i = 0
'''''    Set RSTFILE = New ADODB.Recordset
'''''    RSTFILE.Open "Select * From Pht001", Conn2, adOpenForwardOnly
'''''    Do Until RSTFILE.EOF
'''''        Set RSTFILE2 = New ADODB.Recordset
'''''        RSTFILE2.Open "Select * From CRDTPYMT", DB, adOpenStatic, adLockOptimistic, adCmdText
'''''        RSTFILE2.AddNew
'''''        i = i + 1
'''''        RSTFILE2!TRX_TYPE = "CR"
'''''        RSTFILE2!CR_NO = i
'''''        RSTFILE2!INV_NO = i
'''''        RSTFILE2!INV_DATE = Format(RSTFILE!INVDT, "DD/MM/YYYY")
'''''        RSTFILE2!INV_AMT = RSTFILE!INVAMT
'''''        RSTFILE2!RCPT_AMT = 0
'''''        RSTFILE2!BAL_AMT = RSTFILE!INVAMT
'''''        RSTFILE2!ACT_CODE = "311" & RSTFILE!SUPCD
'''''        RSTFILE2!CHECK_FLAG = "N"
'''''        RSTFILE2!PINV = RSTFILE!INVNO
'''''
'''''        RSTFILE2.Update
'''''        RSTFILE2.Close
'''''        Set RSTFILE2 = Nothing
'''''        RSTFILE.MoveNext
'''''    Loop
'''''    RSTFILE.Close
'''''    Set RSTFILE = Nothing
'''''
'''''
End Sub

Private Function backup_database(Rem_Drive As String)
    Dim SourceFile, DestinationFile, tryagain, result
    Dim strBackupEXT As String
    Dim n As Integer
    
    MDIMAIN.vbalProgressBar1.Visible = True
    MDIMAIN.vbalProgressBar1.Value = 0
    MDIMAIN.vbalProgressBar1.ShowText = True
    MDIMAIN.vbalProgressBar1.Text = "PLEASE WAIT..."
    MousePointer = vbHourglass
    strBackupEXT = "bk" & Format(Format(Date, "ddmmyy"), "000000") & Format(Format(Time, "HHMMSS"), "")
    'Backup Da
    
    
    On Error GoTo handler
    Dim cmd As String
    Screen.MousePointer = vbHourglass
    If Dir(App.Path & "\Backup", vbDirectory) = "" Then MkDir App.Path & "\Backup"
    If Not FileExists(App.Path & "\mysqldump.exe") Then
        Screen.MousePointer = vbNormal
        MsgBox "File not exists", , "EzBiz"
        Exit Function
    End If
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " > " & App.Path & "\Backup\" & strBackupEXT
    Call execCommand(cmd)
    Sleep (300)
    
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " > " & Rem_Drive & strBackupEXT
    Call execCommand(cmd)
    
    err.Clear
    
    Screen.MousePointer = vbNormal
    MDIMAIN.vbalProgressBar1.Text = "Successfully Completed..."
    MsgBox "Back-up complete !!", vbOKOnly, "Back Up!!!!"
    MDIMAIN.vbalProgressBar1.Visible = False
    Exit Function
    
handler:
    Screen.MousePointer = vbNormal
    Select Case err.Number
        Case 70
        MsgBox "Error No. > " & err.Number & " / " & err.Description
        Resume Next
        Case 75
        Resume Next
        Case Else
        MsgBox "Error No. > " & err.Number & " / " & err.Description
    End Select
End Function

Private Sub CMDDUPPURCHASE_Click()
    If MDIMAIN.LBLSHOPRT.Caption = "Y" Then
        frmLWS.Show
        frmLWS.SetFocus
    Else
        If MDIMAIN.lblcategory.Caption = "Y" Then
            frmLW.Show
            frmLW.SetFocus
        Else
            frmLW1.Show
            frmLW1.SetFocus
        End If
    End If
End Sub

Private Sub MNULEND_Click()
    frmLenders.Show
    frmLenders.SetFocus
End Sub

Private Sub CmdPayment_Click()
    FRMPaymntreg.Show
    FRMPaymntreg.SetFocus
End Sub

Private Sub CmdReceipt_Click()
    On Error Resume Next
    FRMRcptReg.Show
    FRMRcptReg.SetFocus
End Sub

Private Sub CmdExp_Click()
    Frmexpense.Show
    Frmexpense.SetFocus
End Sub

Private Sub CmdStaff_Click()
    If MDIMAIN.lblsalary.Caption = "Y" Then
        FRMStaffReg.Show
        FRMStaffReg.SetFocus
    Else
        frmExpenseStaff.Show
        frmExpenseStaff.SetFocus
    End If
End Sub

Private Sub CmdLend_Click()
    FRMLendReg.Show
    FRMLendReg.SetFocus
End Sub

Private Sub CmDIncome_Click()
    frmIncome.Show
    frmIncome.SetFocus
End Sub

Private Sub CmdBook_Click()
    FRMCASHBOOK.Show
    FRMCASHBOOK.SetFocus
End Sub

Private Sub MnuIncomeMast_Click()
    frmIncMast.Show
    frmIncMast.SetFocus
End Sub

Private Function Keydown(key As Integer)
'    Dim dt_from As Date
'    dt_from = "13/04/2021"
'    Dim rstTRXMAST As ADODB.Recordset
    On Error GoTo ERRHAND
'    If Shift = vbCtrlMask Then
        Select Case key
            Case 97, 49
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing
                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                If bill_type_flag = True Then
                    FRMPETTY.Show
                    FRMPETTY.SetFocus
                Else
                    FRMPETTY_TYPE.Show
                    FRMPETTY_TYPE.SetFocus
                End If
            Case 98, 50
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing
                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                
                FRMPETTY1.Show
                FRMPETTY1.SetFocus
            Case 99, 51
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing

                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                
                FRMPETTY2.Show
                FRMPETTY2.SetFocus
        Case 99, 56
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing

                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                
                frmSalesReturnw.Show
                frmSalesReturnw.SetFocus
        Case 99, 57
'                Set rstTRXMAST = New ADODB.Recordset
'                rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
'                If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
'                    MsgBox "Crxdtl.dll cannot be loaded.", vbCritical, "EzBiz"
'                    rstTRXMAST.Close
'                    Set rstTRXMAST = Nothing
'                    Exit Function
'                End If
'                rstTRXMAST.Close
'                Set rstTRXMAST = Nothing

                If exp_flag = True Then
                    'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
                    Call errcodes(Val(lblec.Caption))
                    Exit Function
                End If
                
                FRMPURCHASERETW.Show
                FRMPURCHASERETW.SetFocus
        End Select
'    End If
    Exit Function
    
ERRHAND:
    MsgBox err.Description

End Function

Private Sub TxtDUP_Change()
    If DUPCODE = Trim(TxtDUP.Text) Then
        If exp_flag = True Then
            'MsgBox "Error Code: U1431 Error in fetching data. Probably database currupted.", vbCritical, "EzBiz"
            Exit Sub
        End If
        Call CMDDUPPURCHASE_Click
        TxtDUP.Text = ""
        TxtDUP.Visible = False
    End If
End Sub

Private Sub TxtDUP_GotFocus()
    TxtDUP.SelStart = 0
    TxtDUP.SelLength = Len(TxtDUP.Text)
End Sub


Function GetDataFromURL_APP(strURL, strMethod, strPostData)
  Dim lngTimeout
  Dim strUserAgentString
  Dim intSslErrorIgnoreFlags
  Dim blnEnableRedirects
  Dim blnEnableHttpsToHttpRedirects
  Dim strHostOverride
  Dim strLogin
  Dim strPassword
  Dim strResponseText
  Dim objWinHttp
  
  On Error GoTo ERRHAND
  lngTimeout = 59000
  strUserAgentString = "zen_request/0.1"
  intSslErrorIgnoreFlags = 13056 ' 13056: ignore all err, 0: accept no err
  blnEnableRedirects = True
  blnEnableHttpsToHttpRedirects = True
  strHostOverride = ""
  strLogin = ""
  strPassword = ""
  Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1")
  objWinHttp.setTimeouts lngTimeout, lngTimeout, lngTimeout, lngTimeout
  objWinHttp.Open strMethod, strURL
  If strMethod = "POST" Then
    objWinHttp.setRequestHeader "Content-type", _
      "application/json"

    'objWinHttp.setRequestHeader "Token", _
      strToken

  End If
  If strHostOverride <> "" Then
    objWinHttp.setRequestHeader "Host", strHostOverride
  End If
  objWinHttp.Option(0) = strUserAgentString
  objWinHttp.Option(4) = intSslErrorIgnoreFlags
  objWinHttp.Option(6) = blnEnableRedirects
  objWinHttp.Option(12) = blnEnableHttpsToHttpRedirects
  If (strLogin <> "") And (strPassword <> "") Then
    objWinHttp.SetCredentials strLogin, strPassword, 0
  End If
  On Error Resume Next
  objWinHttp.send (strPostData)
  If err.Number = 0 Then
    If objWinHttp.Status = "200" Then
      GetDataFromURL_APP = objWinHttp.responseText
    Else
      GetDataFromURL_APP = "HTTP " & objWinHttp.Status & " " & _
        objWinHttp.statusText
    End If
  Else
    GetDataFromURL_APP = "Error " & err.Number & " " & err.source & " " & _
      err.Description
  End If
  On Error GoTo 0
  Set objWinHttp = Nothing
  Exit Function
ERRHAND:
    MsgBox err.Description, , "EzBiz"
End Function


Sub SendJSONPOSTRequest()
    
    Dim JSONNAMES As String
    Dim RSTITEMS As ADODB.Recordset
    Set RSTITEMS = New ADODB.Recordset
    
    db.Execute "SELECT JSON_ARRAYAGG(JSON_OBJECT('ITEM_CODE', ITEM_CODE, 'ITEM_NAME', ITEM_NAME)) AS JSONNAMES from ITEMMAST"
    
    
    RSTITEMS.Open "SELECT JSON_ARRAYAGG(JSON_OBJECT('ITEM_CODE', ITEM_CODE)) from ITEMMAST", db
    RSTITEMS.Open "SELECT JSON_ARRAYAGG(JSON_OBJECT('ITEM_CODE', ITEM_CODE)) AS JSONResult FROM ITEMMAST", db
    'db.Execute "Select JSON_ARRAYAGG(JSON_OBJECT('code',ITEM_CODE,'itemname',ITEM_NAME) AS JSONDATAS FROM ITEMMAST"
    Dim objWinHttp As Object
    Set objWinHttp = CreateObject("WinHttp.WinHttpRequest.5.1") ' Create a WinHTTPRequest object

    ' Define the URL to which you want to send the POST request
    Dim URL As String
    URL = "http://192.168.29.182:5400/test"

    ' Set up the request
    objWinHttp.Open "POST", URL, False ' Third parameter: Asynchronous (False for synchronous)

    ' Set request headers to indicate JSON content
    objWinHttp.setRequestHeader "Content-Type", "application/json"
    objWinHttp.setRequestHeader "Accept", "application/json"

    ' Define the JSON data as a string
    Dim jsonData As String
    'jsonData = "{""param1"": ""value1"", ""param2"": ""value2""}" ' Replace with your actual JSON data
    'jsonData = "{""query"": ""Select * from ord_trxfile left join ord_mast on ord_trxfile.ord_no = ord_mast.ord_no where ord_mast.status_flag ='0'""}" ' Replace with your actual JSON data
    jsonData = "{""query"": ""update ord_mast set status_flag = '1' where ord_no = '1'""}" ' Replace with your actual JSON data
    ' Send the JSON POST request
    objWinHttp.send jsonData

    ' Check the response status
    If objWinHttp.Status = 200 Then
        ' Request was successful, and response is available in objWinHTTP.responseText
        Debug.Print "Response: " & objWinHttp.responseText
    Else
        ' Request failed, and you can handle errors here
        Debug.Print "Request failed with status: " & objWinHttp.Status
        Debug.Print "Response Text: " & objWinHttp.responseText
    End If

    ' Clean up
    Set objWinHttp = Nothing
End Sub

