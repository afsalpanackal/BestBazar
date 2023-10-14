Attribute VB_Name = "MDLMEDICINE"
Option Explicit
Declare Function apiCopyFile Lib "kernel32" Alias "CopyFileA" _
(ByVal lpExistingFileName As String, _
ByVal lpNewFileName As String, _
ByVal bFailIfExists As Long) As Long

Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal bytes As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public db As New ADODB.Connection
'Public dbprint As New ADODB.Connection
'Public TEMP As ADODB.Recordset
Public creditbill As Form
Public CALCODE As String
Public DUPCODE As String
Public bill_for As String
Private Const MAX_PATH As Long = 260
Private m_VolName As String
Private m_VolSN As Long
Private m_MaxLen As Long
Private m_Flags As Long
Private m_FileSys As String
Public Cancelbill_flag As Boolean

Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
 
Private Const SYNCHRONIZE       As Long = &H100000
Private Const INFINITE          As Long = &HFFFF

Dim m_Drive As String

Public bill_type_flag As Boolean
Public scheme_option As String
Public system_name As String
Public strConnection As String
Public Rptpath As String
Public DMPrint As String
Public DMPrintA4 As String
Public BarPrint As String
Public dbase1 As String
Public dbase2 As String
Public DBPath As String
Public RstBill_Flag As String
Public exp_flag As Boolean
Public hold_thermal As Boolean

Public customercode As String
Public serveraddress As String
Public remoteflag As Boolean
Public REMOTEAPP As Single
Public MRPDISC_FLAG As String
Public billprinter As String
Public billprinterA5 As String
Public thermalprinter As String
Public barcodeprinter As String
Public Roundflag As Boolean
Public BATCH_DISPLAY As Boolean


Public PC_FLAG As String
Public ZERO_WARN_FLAG As Boolean
Public MINUS_BILL As String
Public SALESLT_FLAG As String
Public BARTEMPLATE As String
'Public BARFORMAT As String

Public D_PRINT As String
Public PRNPETTYFLAG As Boolean

Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Public crxApplication As New CRAXDRT.Application
Public Report As CRAXDRT.Report
Public oRs As ADODB.Recordset
Public CRXFormulaFields As CRAXDRT.FormulaFieldDefinitions
Public CRXFormulaField As CRAXDRT.FormulaFieldDefinition
Public ReportNameVar As String
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Const LOCALE_SDATE                 As Long = &H1D    'date separator
Private Const LOCALE_STIME                 As Long = &H1E    'time separator
Private Const LOCALE_SSHORTDATE            As Long = &H1F    'short date format string
Private Const LOCALE_SLONGDATE             As Long = &H20    'long date format string
Private Const LOCALE_STIMEFORMAT           As Long = &H1003  'time format string
Private Const LOCALE_IDATE                 As Long = &H21    'short date format ordering
Private Const LOCALE_ILDATE                As Long = &H22    'long date format ordering
Private Const LOCALE_ITIME                 As Long = &H23    'time format specifier
 
Private Declare Function SetLocaleInfo& Lib "kernel32" Alias "SetLocaleInfoA" (ByVal _
Locale As Long, ByVal LCType As Long, ByVal lpLCData As String)

Private Declare Function InternetGetConnectedState _
              Lib "wininet.dll" (ByRef lpdwFlags As Long, _
              ByVal dwReserved As Long) As Long


Sub Main()
'    Dim strConn As String
'    Dim strConnect As String
'    Dim DBPwd As String
'    Dim DBPath As String
    

'    Open Rptpath & "cer.txt" For Output As #1 '//Report file CREATION
'    Print #1, GetUniqueCode
'    Close #1
'    Exit Sub


'    If Month(Date) > 9 Or Month(Date) < 5 Then
'        Open "C:\WINDOWS\system32\mwp.lp1" For Output As #1 '//Report file Creation
'        Open "C:\WINDOWS\system32\mwp.lp1" For Output As #1 '//Report file Creation
'        Print #1, ""
'        Close #1
'        End
'        Exit Sub
'    End If
'
'        If FileLen("C:\WINDOWS\system32\appmgmt\S-1-5-21-329068152-1425521274-839522115-1003\wsys.dxd") <> 3 Then GoTo eRRHAND
'         If FileLen("C:\WINDOWS\system32\mwp.lp1") <> 6 Then GoTo eRRHAND
'        If GetUniqueCode <> "791084950" Then GoTo eRRHAND  'Shradha
'        If GetUniqueCode <> "1352116014" Then GoTo eRRHAND 'Surya
        'If GetUniqueCode <> "813765986" Then GoTo errHand 'gEO
'        If GetUniqueCode <> "355072539" Then GoTo eRRHAND 'AMAZONE
'        If GetUniqueCode <> "1598631032" Then GoTo eRRHAND 'Electro
'        If GetUniqueCode <> "707442519" Then GoTo eRRHAND 'SF Sonic
'        If GetUniqueCode <> "739655203" Then GoTo eRRHAND 'Kannattu
'        If GetUniqueCode <> "322694026" Then GoTo eRRHAND 'Real
'        If GetUniqueCode <> "1346214219" Then GoTo eRRHAND 'Muscat
'        If GetUniqueCode <> "1394268811" Then GoTo errHand 'VISHNU D drive
'        If GetUniqueCode <> "1256447070" Then GoTo eRRHAND 'PANDALAM E drive
'        If GetUniqueCode <> "18745649" Then GoTo eRRHAND 'Cherthala F drive
'        If GetUniqueCode <> "845641114" Then GoTo eRRHAND 'Kevina D drive
'        If GetUniqueCode <> "521173863" Then GoTo eRRHAND 'C Sun F drive
        'If GetUniqueCode <> "422390192" And GetUniqueCode <> "1200564647" Then GoTo eRRHAND 'Sreebudha D drive
'        If GetUniqueCode <> "212775474" Then GoTo eRRHAND 'Shaji Sir D drive
'        If GetUniqueCode <> "18745649" Then GoTo eRRHAND 'Initiative  E drive
        'If GetUniqueCode <> "550313660" Then GoTo eRRHAND 'Initiative  E drive
'        If GetUniqueCode <> "950297362" Then GoTo ErrHand 'Palace traders E drive
        'If GetUniqueCode <> "1611126783" And GetUniqueCode <> "1627632539" Then GoTo eRRHAND 'Bismi D drive
'        If GetUniqueCode <> "1902784888" Then GoTo eRRHAND 'VP Stores
'        If GetUniqueCode <> "774754286" Then GoTo eRRHAND 'VP Stores
        'If GetUniqueCode <> "1959685209" Then GoTo eRRHAND 'Riti Trading
        'If GetUniqueCode <> "1639577482" Then GoTo errHand 'Petz
'        If GetUniqueCode <> "1828610313" Then GoTo ERRHAND 'TU AGENCIES
        'If GetUniqueCode <> "942043983" Then GoTo eRRHAND 'CC Traders
        'If GetUniqueCode <> "1246050505" Then GoTo eRRHAND 'Vadakuzhy Traders
        'If GetUniqueCode <> "1238340068" And GetUniqueCode <> "599112474" Then GoTo eRRHAND 'AR STEELS
        'If GetUniqueCode <> "485948363" Then GoTo ERRHAND 'CITY WHEELS
        'If GetUniqueCode <> "1791263069" And GetUniqueCode <> "8474593" Then GoTo eRRHAND 'Real
        'If GetUniqueCode <> "1624680873" And GetUniqueCode <> "1009067786" Then GoTo eRRHAND 'Thoppil
        'If GetUniqueCode <> "301496399" Then GoTo eRRHAND 'vELIYIL
        'If GeftUniqueCode <> "389073820" Then GoTo eRRHAND 'Modern
        'If GetUniqueCode <> "340476196" Then GoTo eRRHAND 'Aksa
        'If GetUniqueCode <> "52535345" Then GoTo eRRHAND 'Kebi
        'If GetUniqueCode <> "875518048" Then GoTo eRRHAND 'Robin
        'If GetUniqueCode <> "841114577" Then GoTo eRRHAND 'GENERAL
        'If GetUniqueCode <> "2041237382" Then GoTo eRRHAND 'Krishna
        'If GetUniqueCode <> "732426451" Then GoTo eRRHAND 'Bharath Gum
        'If GetUniqueCode <> "301496399" Then GoTo eRRHAND 'SELF
        'If GetUniqueCode <> "457481987" Then GoTo eRRHAND 'mANGANTHANAM
        'If GetUniqueCode <> "237638351" Then GoTo eRRHAND 'AH Enterprises
        'If GetUniqueCode <> "1801570118" Then GoTo errHand 'Venice Marketing
        'If GetUniqueCode <> "2109330768" Then GoTo eRRHAND 'RENJITH MANKOMPU
        'If GetUniqueCode <> "926059395" Then GoTo eRRHAND 'HANNA MARKETING
        'If GetUniqueCode <> "1674041965" Then GoTo eRRHAND 'Reni Adoor
        'If GetUniqueCode <> "505604364" Then GoTo eRRHAND 'JACKS
        'If GetUniqueCode <> "1927256233" Then GoTo ErrHand 'EZUNOOTTIL
        'If GetUniqueCode <> "2098527197" Then GoTo eRRHAND 'hAPPY rOOMS
        'If GetUniqueCode <> "1263583449" Then GoTo eRRHAND 'sun
        'If GetUniqueCode <> "1667919129" Then GoTo eRRHAND 'Pavithran Provision Thiruvambady
        'If GetUniqueCode <> "727953592" Then GoTo eRRHAND 'Royal Associates, Koottuveli
        'If GetUniqueCode <> "1646307026" Then GoTo eRRHAND 'CM Electricals, Valiyakulam
        'If GetUniqueCode <> "15313424" Then GoTo eRRHAND 'Milma Shoppee
        'If GetUniqueCode <> "661852731" Then GoTo eRRHAND 'KOCHIN TRADERS
        'If GetUniqueCode <> "617501942" Then GoTo errHand 'SHAN TOOLS
        'If GetUniqueCode <> "2075798045" And GetUniqueCode <> "729014108" Then GoTo eRRHAND 'ThreeStar
        'If GetUniqueCode <> "1577112679" Then GoTo eRRHAND  'Falcon Associate
        'If GetUniqueCode <> "812567791" Then GoTo eRRHAND  'Falcon Tiles1
        'If GetUniqueCode <> "1482138994" Then GoTo eRRHAND  'Aisha
        'If GetUniqueCode <> "203727769" And GetUniqueCode <> "1044911181" Then GoTo eRRhAND  'AB AGENCIES
        'If GetUniqueCode <> "1129152818" Then GoTo eRRHAND  'Aisha
        'If GetUniqueCode <> "1482138994" Then GoTo errHand  'Aisha
        'If GetUniqueCode <> "26782297" Then GoTo errHand  'Aisha
        'If GetUniqueCode <> "286930845" And GetUniqueCode <> "1445923169" Then GoTo errHand  'AM STEELS
        'If GetUniqueCode <> "670182758" Then GoTo errHand  'RR DENTALS
        'If GetUniqueCode <> "2128428739" Then GoTo eRRHAND  'MANGALASSERY TRADERS
        'If GetUniqueCode <> "1475509127" Then GoTo errHand  'MANGALASSERY TRADERS
        'If GetUniqueCode <> "1940992988" Then GoTo eRRHAND  'ASHWIN TRADERS
        'If GetUniqueCode <> "1246050505" And GetUniqueCode <> "1286435747" Then GoTo eRRHAND    'Sun Technolgies
        'If GetUniqueCode <> "4406466" Then GoTo eRRHAND  'Jetty Agencies
        'If GetUniqueCode <> "432897505" Then GoTo eRRHAND  'Amni Tyres
        'If GetUniqueCode <> "1680960773" Then GoTo eRRHAND  'PhotoLine
        'If GetUniqueCode <> "231288830" And GetUniqueCode <> "392535037" Then GoTo errHand  'Real & Ideal
        'If GetUniqueCode <> "1059775097" Then GoTo errHand 'sELF E drive
        'If GetUniqueCode <> "517376339" Then GoTo eRRHAND  'aKSHAYA kATTOOR
        'If GetUniqueCode <> "2113334822" Then GoTo eRRHAND  'CVS Enterprises
        'If GetUniqueCode <> "1195565218" Then GoTo eRRHAND  'AR CAMP
 ''MsgBox GetUniqueCode
'        If GetUniqueCode <> "521173863" And GetUniqueCode <> "889572787" Then GoTo ErrHand 'C Sun
''
''    'Exit Sub
    'MsgBox "Check Authorisation"
    'Amazone D Drive
    'C-SUN - e dRIVE
        
    Dim LCID As Long
    Call SetLocaleInfo(LCID, LOCALE_SSHORTDATE, CStr("dd/MM/yyyy"))
    
    Dim strConn As String
    Dim ObjFso
    Dim StrFileName
    Dim ObjFile
    Dim sitem As String
    If FileExists(App.Path & "\EX") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\EX")  'Reading from the file
        Rptpath = ObjFile.ReadLine
    Else
        sitem = UCase(InputBox("Path?", ""))
        If Trim(sitem) = "" Then
            MsgBox "Invalid Path", vbOKOnly, ""
            End
        End If
        StrFileName = App.Path & "\EX"
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.CreateTextFile(StrFileName)
        ObjFile.Write sitem
        Rptpath = sitem
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    If Right(Rptpath, 1) <> "\" Then
        Rptpath = Rptpath & "\"
    End If
    
    If FileExists(App.Path & "\db") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\db")  'Reading from the file
        dbase1 = ObjFile.ReadLine
    Else
        sitem = UCase(InputBox("db?", ""))
        If Trim(sitem) = "" Then
            MsgBox "db not yet prepared", vbOKOnly, ""
            End
        End If
        StrFileName = App.Path & "\db"
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.CreateTextFile(StrFileName)
        ObjFile.Write sitem
        dbase1 = sitem
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    If FileExists(App.Path & "\db2") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\db2")  'Reading from the file
        dbase2 = ObjFile.ReadLine
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    If FileExists(App.Path & "\BillPrint") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\BillPrint")  'Reading from the file
        billprinter = ObjFile.ReadLine
        On Error Resume Next
        billprinterA5 = ObjFile.ReadLine
        thermalprinter = ObjFile.ReadLine
        barcodeprinter = ObjFile.ReadLine
        If barcodeprinter = "" Then barcodeprinter = billprinter
        
        err.Clear
        On Error GoTo ERRHAND
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
    
    
'    Set db = New ADODB.Connection
'    '"\\192.168.1.3\data (d)\dbase1
'    DBPwd = "###DATABASE%%%RET"
'
'    If FileLen(DBPath) Then
'
'    End If
'    FrmRestore.Show
'    Exit Sub
    Dim strCnn As String
    Dim sql As String
    Dim MD5 As New clsMD5
    Dim ACT_KEY1, ACT_KEY2 As String
    ACT_KEY1 = Val(GetUniqueCode) * 555
    
    ACT_KEY2 = UCase(MD5.DigestStrToHexStr(ACT_KEY1))
    ACT_KEY2 = ACT_KEY2 & UCase(MD5.DigestStrToHexStr(ACT_KEY2))
    ACT_KEY2 = Mid(ACT_KEY2, 24, 10) & Mid(ACT_KEY2, 1, 5)
    
    'strConnect = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & DBPath & ";Jet OLEDB:Database Password=" & DBPwd
    
    DBPath = "localhost"
    
    Dim dwLen As Long
        'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    system_name = String(dwLen, "X")
    'Get the computer name
    GetComputerName system_name, dwLen
    'get only the actual data
    system_name = Left(system_name, dwLen)
    
    
    
    If FileExists(App.Path & "\Client.txt") Then
        Set ObjFso = CreateObject("Scripting.FileSystemObject")  'Opening the file in READ mode
        Set ObjFile = ObjFso.OpenTextFile(App.Path & "\Client.txt")  'Reading from the file
        DBPath = ObjFile.ReadLine
    End If
    Set ObjFso = Nothing
    Set ObjFile = Nothing
       
    On Error GoTo ERRHAND
    
    'DBPath = "https://gold.primecrown.net:2083/"
    'strConn = "Driver={MySQL ODBC 3.51 Driver};Server=" & DBPath & ";Port=3306;Database=ezbizco_data;User=ezbizco; Password=D]&$9xUOmIS{; Option=2;"
    
    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase1 & ";User=root; Password=###%%database%%###ret; Option=2;"
    db.Open strConn
    db.CursorLocation = adUseClient
    
'    If MsgBox("Do yu want to regenerate db?", vbYesNo + vbDefaultButton2, "EzBiz.....") = vbYes Then
'        db.Execute "CHECK TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
'        db.Execute "CHECK TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
'        db.Execute "CHECK TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
'        db.Execute "CHECK TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
'        db.Execute "CHECK TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
'        db.Execute "CHECK TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
'        db.Execute "CHECK TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
'        db.Execute "CHECK TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
'        db.Execute "CHECK TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
'        db.Execute "CHECK TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
'
'        db.Execute "OPTIMIZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
'        db.Execute "OPTIMIZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
'        db.Execute "OPTIMIZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
'        db.Execute "OPTIMIZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
'        db.Execute "OPTIMIZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
'        db.Execute "OPTIMIZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
'        db.Execute "OPTIMIZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
'        db.Execute "OPTIMIZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
'        db.Execute "OPTIMIZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
'        db.Execute "OPTIMIZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
'
'        db.Execute "REPAIR TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
'        db.Execute "REPAIR TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
'        db.Execute "REPAIR TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
'        db.Execute "REPAIR TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
'        db.Execute "REPAIR TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
'        db.Execute "REPAIR TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
'        db.Execute "REPAIR TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
'        db.Execute "REPAIR TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
'        db.Execute "REPAIR TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
'        db.Execute "REPAIR TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
'
'        db.Execute "ANALYZE TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
'        db.Execute "ANALYZE TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
'        db.Execute "ANALYZE TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
'        db.Execute "ANALYZE TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
'        db.Execute "ANALYZE TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
'        db.Execute "ANALYZE TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
'        db.Execute "ANALYZE TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
'        db.Execute "ANALYZE TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
'        db.Execute "ANALYZE TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
'        db.Execute "ANALYZE TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
'
'        db.Execute "FLUSH TABLE `actmast`, `address_book`, `arealist`, `atrxfile`, `atrxsub`"
'        db.Execute "FLUSH TABLE `bankcode`, `bankletters`, `bank_trx`, `barprint`, `billdetails`, `bonusmast`, `bookfile`, `cancinv`, `cashatrxfile`"
'        db.Execute "FLUSH TABLE `category`, `chqmast`, `compinfo`, `cont_mast`, `crdtpymt`, `custmast`, `custtrxfile`, `damaged`, `DAMAGE_MAST`, `dbtpymt`"
'        db.Execute "FLUSH TABLE `de_ret_details`, `expiry`, `explist`, `expsort`, `fqtylist`, `gift`, `itemmast`, `manufact`, `moleculelink`, `molecules`"
'        db.Execute "FLUSH TABLE `nonrcvd`, `ordermast`, `ordertrxfile`, `ordissue`, `ordsub`, `password`, `passwords`, `pomast`, `posub`"
'        db.Execute "FLUSH TABLE `pricetable`, `prodlink`, `purcahsereturn`, `purch_return`, `qtnmast`, `qtnsub`, `reorder`, `replcn`, `returnmast`, `roomtrxfile`"
'        db.Execute "FLUSH TABLE `rtrxfile`, `salereturn`, `salesledger`, `salesman`, `salesreg`, `seldist`, `service_stk`, `slip_reg`, `srtrxfile`, `stockreport`"
'        db.Execute "FLUSH TABLE `tbletrxfile`, `tempcn`, `tempstk`, `temptrxfile`, `tmporderlist`, `transmast`, `transsub`, `trnxrcpt`"
'        db.Execute "FLUSH TABLE `trxexpense`, `trxexpmast`, `trxexp_mast`, `trxfile`, `trxfileexp`, `trxfile_exp`, `trxfile_formula`, `trxfile_sp`"
'        db.Execute "FLUSH TABLE `trxincmast`, `trxincome`, `trxmast`, `trxmast_sp`, `trxsub`, `users`, `vanstock`, `war_list`, `war_trxfile`, `war_trxns`"
'
'        On Error Resume Next
'        db.Execute "CHECK TABLE `astmast`, `astrxfile`, `astrxmast`, "
'        db.Execute "OPTIMIZE TABLE `astmast`, `astrxfile`, `astrxmast`, "
'        db.Execute "REPAIR TABLE `astmast`, `astrxfile`, `astrxmast`, "
'
'        db.Execute "CHECK TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
'        db.Execute "OPTIMIZE TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
'        db.Execute "REPAIR TABLE `trxformulasub`, `trxformulamast`, `trnxroom`, `trnxtable`"
'
'        On Error GoTo ErrHand
'    End If
'    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=tempdb;User=root; Password=###%%database%%###ret; Option=2;"
'    dbprint.Open strConn
'    dbprint.CursorLocation = adUseClient
'
    
    strConnection = "Provider=MSDASQL;Driver={MySQL ODBC 5.1 Driver};Server=" & DBPath & ";Port=3306;Database=" & dbase1 & ";User=root; Password=###%%database%%###ret"
    'strConnection = "Provider=MSDASQL;" & _
    "SERVER=" & DBPath & ";PORT=3306;DATABASE=" & dbase1 & ";" & _
    "USER=root;PASSWORD=###%%database%%###ret;"
    
    Dim TRXFILE As ADODB.Recordset
    Dim TRXFILE2 As ADODB.Recordset
    Set TRXFILE = New ADODB.Recordset
    sql = "select * from act_ky WHERE ACT_CODE= '" & ACT_KEY2 & "'"
    TRXFILE.Open sql, db, adOpenKeyset, adLockPessimistic
    If TRXFILE.BOF And TRXFILE.EOF Then
        FrmKey.Show
        FrmKey.SetFocus
        FrmKey.lblINSID.Caption = ACT_KEY1
        TRXFILE.Close
        Set TRXFILE = Nothing
        Exit Sub
    Else
        Dim exp_days As Integer
        Dim dt_from As Date
        Dim dt_to As Date
        Dim dt_exp As Date
        
        exp_days = 0
        If Not IsNull(TRXFILE!actky4) Then
            If IsDate(EncryptString(TRXFILE!actky4, "ezkeys")) Then dt_exp = EncryptString(TRXFILE!actky4, "ezkeys")
        End If
        If IsDate(EncryptString(TRXFILE!actky2, "ezbizkeys")) Then
            
            If Not TRXFILE!actky3 = Mid(UCase(MD5.DigestStrToHexStr(TRXFILE!ACT_CODE & EncryptString(TRXFILE!actky2, "keyezbiz"))), 1, 20) Then
                FrmKey.Show
                FrmKey.SetFocus
                FrmKey.lblINSID.Caption = ACT_KEY1
                TRXFILE.Close
                Set TRXFILE = Nothing
                Exit Sub
            End If
            
            dt_from = EncryptString(TRXFILE!actky2, "ezbizkeys")
            dt_to = DateAdd("d", 16, dt_from)
            Dim rstTRXMAST As ADODB.Recordset
            Dim rstTRXMAST2 As ADODB.Recordset
            
            If DateDiff("d", Date, dt_from) <= 0 And DateDiff("d", Date, dt_from) > -15 Then
                If MsgBox("Annual Service Package expired!! Kindly renew your Annual Service Package." & Chr(13) & "You will be on a grace period. Do you want to continue? Press No to Activate.", vbYesNo + vbDefaultButton2, "EzBiz Activation") = vbYes Then
                    Set rstTRXMAST = New ADODB.Recordset
                    rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
                    If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                        MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                        rstTRXMAST.Close
                        Set rstTRXMAST = Nothing
                        FrmKey.Show
                        FrmKey.SetFocus
                        FrmKey.lblINSID.Caption = ACT_KEY1
                        Exit Sub
                    Else
                        Set rstTRXMAST2 = New ADODB.Recordset
                        rstTRXMAST2.Open "SELECT * From rtrxfile WHERE VCH_DATE >= '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
                        If Not (rstTRXMAST2.EOF And rstTRXMAST2.BOF) Then
                            MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                            rstTRXMAST2.Close
                            Set rstTRXMAST2 = Nothing
                            rstTRXMAST.Close
                            Set rstTRXMAST = Nothing
                            FrmKey.Show
                            FrmKey.SetFocus
                            FrmKey.lblINSID.Caption = ACT_KEY1
                            Exit Sub
                        End If
                        rstTRXMAST2.Close
                        Set rstTRXMAST2 = Nothing
                        If DateDiff("d", Date, dt_exp) <= 0 Then
                            exp_flag = True
                            db.Execute "Update COMPINFO set EC = " & Day(Date) & " where COMP_CODE = '001' AND EC =0"
                        'Else
                        '    db.Execute "Update COMPINFO set EC =0"
                        End If
                        rstTRXMAST.Close
                        Set rstTRXMAST = Nothing
                        GoTo SKIP
                    End If
                Else
                    FrmKey.Show
                    FrmKey.SetFocus
                    FrmKey.lblINSID.Caption = ACT_KEY1
                    Exit Sub
                End If
            End If
            
            dt_to = DateAdd("d", 16, dt_from)
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                rstTRXMAST.Close
                Set rstTRXMAST = Nothing
                MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                
                FrmKey.Show
                FrmKey.SetFocus
                FrmKey.lblINSID.Caption = ACT_KEY1
                TRXFILE.Close
                Set TRXFILE = Nothing
                Exit Sub
            Else
                Set rstTRXMAST2 = New ADODB.Recordset
                rstTRXMAST2.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
                If Not (rstTRXMAST2.EOF And rstTRXMAST2.BOF) Then
                    rstTRXMAST.Close
                    Set rstTRXMAST = Nothing
                
                    rstTRXMAST2.Close
                    Set rstTRXMAST2 = Nothing
                    MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                    
                    FrmKey.Show
                    FrmKey.SetFocus
                    FrmKey.lblINSID.Caption = ACT_KEY1
                    TRXFILE.Close
                    Set TRXFILE = Nothing
                    Exit Sub
                End If
                rstTRXMAST2.Close
                Set rstTRXMAST2 = Nothing
                
                If IsDate(dt_exp) Then
                    Set rstTRXMAST2 = New ADODB.Recordset
                    rstTRXMAST2.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_exp, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
                    If Not (rstTRXMAST2.EOF And rstTRXMAST2.BOF) Then
                        exp_flag = True
                        db.Execute "Update COMPINFO set EC = " & Day(Date) & " where COMP_CODE = '001' AND EC =0"
                    Else
                        exp_flag = False
'                        db.Execute "Update COMPINFO set EC =0"
                    End If
                    rstTRXMAST2.Close
                    Set rstTRXMAST2 = Nothing
                End If
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            If DateDiff("d", Date, dt_from) <= 0 Then
                MsgBox "Annual Service Package expired!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                
                FrmKey.Show
                FrmKey.SetFocus
                FrmKey.lblINSID.Caption = ACT_KEY1
                TRXFILE.Close
                Set TRXFILE = Nothing
                Exit Sub
            End If
            
            If DateDiff("d", Date, dt_from) <= 30 Then
                MsgBox "Annual Service Package expires within " & DateDiff("d", Date, dt_from) & " days!! Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
            End If
            
            Set rstTRXMAST = New ADODB.Recordset
            rstTRXMAST.Open "SELECT * From TRXMAST WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' and VCH_DATE < '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
            If Not (rstTRXMAST.EOF And rstTRXMAST.BOF) Then
                MsgBox "You are on Grace Period. Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
            Else
                Set rstTRXMAST2 = New ADODB.Recordset
                rstTRXMAST2.Open "SELECT * From rtrxfile WHERE VCH_DATE >= '" & Format(dt_from, "yyyy/mm/dd") & "' and VCH_DATE < '" & Format(dt_to, "yyyy/mm/dd") & "' ", db, adOpenStatic, adLockReadOnly
                If Not (rstTRXMAST2.EOF And rstTRXMAST2.BOF) Then
                    MsgBox "You are on Grace Period. Please renew your Annual Service Package", vbOKOnly, "EzBiz Activation"
                End If
                rstTRXMAST2.Close
                Set rstTRXMAST2 = Nothing
            End If
            rstTRXMAST.Close
            Set rstTRXMAST = Nothing
            
            If DateDiff("d", Date, dt_exp) <= 0 Then
                'exp_flag = True
                'db.Execute "Update COMPINFO set EC = " & Day(Date) & " where COMP_CODE = '001' AND EC =0"
            End If
        Else
        
            FrmKey.Show
            FrmKey.SetFocus
            FrmKey.lblINSID.Caption = ACT_KEY1
            TRXFILE.Close
            Set TRXFILE = Nothing
            Exit Sub
        End If
    End If
    TRXFILE.Close
    Set TRXFILE = Nothing
    'bill_for = "0000"
SKIP:
    Set TRXFILE = New ADODB.Recordset
    TRXFILE.Open "SELECT CLCODE FROM COMPINFO WHERE COMP_CODE = '001' ORDER BY FIN_YR DESC", db, adOpenStatic, adLockReadOnly
    If Not (TRXFILE.EOF And TRXFILE.BOF) Then
        CALCODE = IIf(IsNull(TRXFILE!CLCODE), "", TRXFILE!CLCODE)
    End If
    TRXFILE.Close
    Set TRXFILE = Nothing

    If Val(CALCODE) = 0 Then
        frmLogin.Show
    Else
        frmCalculator.Show
    End If
    
    Exit Sub
    
ERRHAND:
    MsgBox err.Description
    'MsgBox "UNAUTHORISED CERTIFICATE", vbCritical, "CERTIFICATE"
    End
End Sub

Public Sub cetre(myForm As Form)

'Place form in center of screen
myForm.Left = (Screen.Width - myForm.Width) / 2
myForm.Top = (Screen.Height - myForm.Height) / 2

End Sub

Public Function AlignLeft(vStr As String, vSpace As Integer) As String
    If Len((vStr)) > vSpace Then '//if the string length is greater than the space you mention
        AlignLeft = Left(vStr, vSpace)  '&"..."
        Exit Function
    End If
    
    AlignLeft = vStr & Space(vSpace - Len(Trim(vStr)))
End Function

Public Function AlignRight(vNumber As String, vSpace As Integer) As String
    AlignRight = Space(vSpace - Len(Trim(vNumber))) & vNumber
End Function

Public Function RepeatString(vStr As String, vSpace As Integer) As String

    Dim x As Integer
    
    For x = 1 To vSpace
        RepeatString = RepeatString & vStr
    Next x
End Function

Public Function GetUniqueCode(Optional ByVal CurDrv As String = "C:\") As Long

    m_VolName = Space(MAX_PATH)
    m_FileSys = Space(MAX_PATH)
    m_Drive = "C:\"
'///// make call, and get Drive Volume Serial Number /////
    If GetVolumeInformation(m_Drive, m_VolName, MAX_PATH, m_VolSN, m_MaxLen, m_Flags, m_FileSys, MAX_PATH) Then
        GetUniqueCode = m_VolSN
    Else
        GetUniqueCode = 0
    End If

'///// Get rid of - (dashes) /////
    GetUniqueCode = Replace(GetUniqueCode, "-", "")

End Function

Public Sub GENERATEREPORT()
    
    frmreport.CRViewer91.ReportSource = Report
    frmreport.ViewReport
    frmreport.Show
    Set CRXFormulaFields = Nothing
    Set CRXFormulaField = Nothing
    Set crxApplication = Nothing
    'Set Report = Nothing
    MDIMAIN.Enabled = False
   Exit Sub

ErrorTrap:
    MsgBox "Error Number: " & err.Number & vbCrLf & err.Description & vbCrLf & vbCrLf & "Debug Information:" & vbCrLf & _
        "ProjectName.FormNamel.cmdSelect_Click" & IIf(Erl > 0, "." & Erl, ""), vbCritical, "Error Occurred"

End Sub

Public Function Words_1_all(ByVal Num As Currency) As String
Dim power_value(1 To 5) As Currency
Dim power_name(1 To 5) As String
Dim digits As Integer
Dim result As String
Dim i As Long

    ' Initialize the power names and values.
    power_name(1) = "trillion": power_value(1) = 1000000000000#
    power_name(2) = "billion":  power_value(2) = 1000000000
    power_name(3) = "million":  power_value(3) = 1000000
    power_name(4) = "thousand": power_value(4) = 1000
    power_name(5) = "":         power_value(5) = 1

    For i = 1 To 5
        ' See if we have digits in this range.
        If Num >= power_value(i) Then
            ' Get the digits.
            digits = Int(Num / power_value(i))

            ' Add the digits to the result.
            If Len(result) > 0 Then result = result & ", "
            result = result & _
                Words_1_999(digits) & _
                " " & power_name(i)

            ' Get the number without these digits.
            Num = Num - digits * power_value(i)
        End If
    Next i

    Words_1_all = Trim$(result)
End Function

 'Return words for this value between 1 and 999.
Private Function Words_1_999(ByVal Num As Integer) As String
Dim hundreds As Integer
Dim remainder As Integer
Dim result As String

    hundreds = Num \ 100
    remainder = Num - hundreds * 100

    If hundreds > 0 Then
        result = Words_1_19(hundreds) & " hundred "
    End If

    If remainder > 0 Then
        result = result & Words_1_99(remainder)
    End If

    Words_1_999 = Trim$(result)
End Function
' Return a word for this value between 1 and 99.
Private Function Words_1_99(ByVal Num As Integer) As String
Dim result As String
Dim tens As Integer

    tens = Num \ 10

    If tens <= 1 Then
        ' 1 <= num <= 19
        result = result & " " & Words_1_19(Num)
    Else
        ' 20 <= num
        ' Get the tens digit word.
        Select Case tens
            Case 2
                result = "Twenty"
            Case 3
                result = "Thirty"
            Case 4
                result = "Forty"
            Case 5
                result = "Fifty"
            Case 6
                result = "Sixty"
            Case 7
                result = "Seventy"
            Case 8
                result = "Eighty"
            Case 9
                result = "Ninety"
        End Select

        ' Add the ones digit number.
        result = result & " " & Words_1_19(Num - tens * 10)
    End If

    Words_1_99 = Trim$(result)
End Function
' Return a word for this value between 1 and 19.
Private Function Words_1_19(ByVal Num As Integer) As String
    Select Case Num
        Case 1
            Words_1_19 = "One"
        Case 2
            Words_1_19 = "Two"
        Case 3
            Words_1_19 = "Three"
        Case 4
            Words_1_19 = "Four"
        Case 5
            Words_1_19 = "Five"
        Case 6
            Words_1_19 = "Six"
        Case 7
            Words_1_19 = "Seven"
        Case 8
            Words_1_19 = "Eight"
        Case 9
            Words_1_19 = "Nine"
        Case 10
            Words_1_19 = "Ten"
        Case 11
            Words_1_19 = "Eleven"
        Case 12
            Words_1_19 = "Twelve"
        Case 13
            Words_1_19 = "Thirteen"
        Case 14
            Words_1_19 = "Fourteen"
        Case 15
            Words_1_19 = "Fifteen"
        Case 16
            Words_1_19 = "Sixteen"
        Case 17
            Words_1_19 = "Seventeen"
        Case 18
            Words_1_19 = "Eightteen"
        Case 19
            Words_1_19 = "Nineteen"
    End Select
End Function

Public Function IsFormLoaded(fForm As Form) As Boolean
    On Error GoTo Err_Proc
    
    Dim x As Integer
    
    For x = 0 To Forms.COUNT - 1
    If (Forms(x) Is fForm) Then
    IsFormLoaded = True
    Exit Function
    End If
    Next x
    
    IsFormLoaded = False
    
Exit_Proc:
    Exit Function
    
Err_Proc:
    MsgBox err.Description, vbOKOnly, "EzBiz"
    Resume Exit_Proc

End Function

Public Function Clipboard_fn(KeyCode As Integer, Shift As Integer, m_textbox As TextBox)
    'Ctrl + A
    If KeyCode = 65 And Shift = 2 Then '
        m_textbox.SelStart = 0
        m_textbox.SelLength = Len(m_textbox.Text)
        'KeyAscii = 0
    End If
    'Ctrl + V
    If KeyCode = 86 And Shift = vbCtrlMask Then
        m_textbox.Text = Clipboard.GetText
        m_textbox.SelStart = Len(m_textbox.Text)
    End If
    If KeyCode = 67 And Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetText m_textbox.SelText
    End If 'Ctrl + c
    If KeyCode = 88 And Shift = vbCtrlMask Then
        Clipboard.Clear
        Clipboard.SetText m_textbox.SelText
        m_textbox.Text = ""
    End If 'Ctrl + X
End Function

Public Function FileExists(ByVal sFileName As String) As Boolean
    Dim intReturn As Integer

    On Error GoTo FileExists_Error
    intReturn = GetAttr(sFileName)
    FileExists = True
    
Exit Function
FileExists_Error:
    FileExists = False
End Function

Function EncryptString(ByVal Text As String, ByVal Password As String) As String
    Dim passLen As Long
    Dim i As Long
    Dim passChr As Integer
    Dim passNdx As Long
    
    passLen = Len(Password)
    ' null passwords are invalid
    If passLen = 0 Then err.Raise 5
    
    ' move password chars into an array of Integers to speed up code
    ReDim passChars(0 To passLen - 1) As Integer
    CopyMemory passChars(0), ByVal StrPtr(Password), passLen * 2
    
    ' this simple algorithm XORs each character of the string
    ' with a character of the password, but also modifies the
    ' password while it goes, to hide obvious patterns in the
    ' result string
    For i = 1 To Len(Text)
        ' get the next char in the password
        passChr = passChars(passNdx)
        ' encrypt one character in the string
        Mid$(Text, i, 1) = Chr$(Asc(Mid$(Text, i, 1)) Xor passChr)
        ' modify the character in the password (avoid overflow)
        passChars(passNdx) = (passChr + 17) And 255
        ' prepare to use next char in the password
        passNdx = (passNdx + 1) Mod passLen
    Next

    EncryptString = Text
    
End Function

Function Decode_Cost(ByVal CostVal As String) As String
    Dim RSTCOMPANY As ADODB.Recordset
    Dim i, n As Integer
    Dim Final_Result As String
    
    On Error GoTo ERRHAND
    CostVal = Int(CostVal)
    Final_Result = ""
    Set RSTCOMPANY = New ADODB.Recordset
    RSTCOMPANY.Open "SELECT * FROM ccode ", db, adOpenStatic, adLockReadOnly
    If Not (RSTCOMPANY.EOF And RSTCOMPANY.BOF) Then
        For i = 1 To Len(CostVal)
            n = Mid(CostVal, i, 1)
            Select Case n
                Case 0
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC0), "", RSTCOMPANY!CC0)
                Case 1
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC1), "", RSTCOMPANY!CC1)
                Case 2
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC2), "", RSTCOMPANY!CC2)
                Case 3
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC3), "", RSTCOMPANY!CC3)
                Case 4
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC4), "", RSTCOMPANY!CC4)
                Case 5
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC5), "", RSTCOMPANY!CC5)
                Case 6
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC6), "", RSTCOMPANY!CC6)
                Case 7
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC7), "", RSTCOMPANY!CC7)
                Case 8
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC8), "", RSTCOMPANY!CC8)
                Case 9
                    Final_Result = Final_Result & IIf(IsNull(RSTCOMPANY!CC9), "", RSTCOMPANY!CC9)
            End Select
        Next i
    End If
    RSTCOMPANY.Close
    Set RSTCOMPANY = Nothing
    
    Decode_Cost = Final_Result
    Exit Function
ERRHAND:
    MsgBox err.Description, , "EzBiz"
    
End Function

Public Function export_db()
'''
'''    If Trim(customercode) = "" Then Exit Function
'''    Dim URL As String
'''    Dim p As Object
'''
'''    On Error GoTo ErrHand
'''    Screen.MousePointer = vbHourglass
'''    Dim http As Object
'''    Set http = CreateObject("WinHttp.WinHttprequest.5.1")
'''    URL = "http://www.ezbiz.co.in/ezbiz/orders.php"
'''    http.Open "Get", URL, False
'''    http.send
'''
    'Set p = JSON.parse(http.responseText)
'''    Dim I As Long
'''    I = 1
''''    For i = 1 To p.COUNT
''''        MsgBox "Order N0. " & p.Item(i).Item("ORDER_NO")
''''        MsgBox "User ID. " & p.Item(i).Item("USER_ID")
''''        MsgBox "Customer Code. " & p.Item(i).Item("ACT_CODE")
''''        MsgBox "Customer Name. " & p.Item(i).Item("ACT_NAME")
''''        MsgBox "Line No. " & p.Item(i).Item("LN_NO")
''''        MsgBox "Item Code. " & p.Item(i).Item("ITEM_CODE")
''''        MsgBox "Item Name. " & p.Item(i).Item("ITEM_NAME")
''''        MsgBox "Unit Price. " & p.Item(i).Item("U_PRICE")
''''        MsgBox "Qty. " & p.Item(i).Item("QTY")
''''        MsgBox "Rate. " & p.Item(i).Item("T_PRICE")
''''    Next i
''''
'''    Dim RSTTRXFILE As ADODB.Recordset
'''    Dim RSTordtrxfile1 As ADODB.Recordset
'''    Dim RSTordtrxfile As ADODB.Recordset
'''    Dim CUSTTYPE As Integer
'''    Dim BILL_NO As Long
'''
'''    'ord_no INT NOT NULL, act_code VARCHAR(25) NULL, act_name VARCHAR(100) NULL, act_address Text(200) NULL, act_phone TEXT(15), C_USER_ID TEXT(4), C_USER_NAME TEXT(50), C_USER_DATE DATE, M_USER_ID TEXT(4), M_USER_NAME TEXT(50), M_USER_DATE DATE, PRIMARY KEY (ord_no)) ENGINE = MyISAM"
'''
'''    db.Execute "delete from orders"
'''    Set RSTTRXFILE = New ADODB.Recordset
'''    RSTTRXFILE.Open "Select * From orders", db, adOpenStatic, adLockOptimistic, adCmdText
'''    'line_no INT NOT NULL, item_code VARCHAR(25) NULL, item_name VARCHAR(200) NULL, item_uprice double null, item_qty double null, item_tprice double null, PRIMARY KEY (ord_no, COMP_CODE, ACT_CODE, line_no)) ENGINE = MyISAM"
'''    For I = 1 To p.COUNT
'''        RSTTRXFILE.AddNew
'''        RSTTRXFILE!ord_no = p.Item(I).Item("ORDER_NO")
'''        RSTTRXFILE!USER_ID = p.Item(I).Item("USER_ID")
'''        RSTTRXFILE!COMP_CODE = p.Item(I).Item("COMP_CODE")
'''        RSTTRXFILE!ACT_CODE = p.Item(I).Item("ACT_CODE")
'''        RSTTRXFILE!ACT_NAME = p.Item(I).Item("ACT_NAME")
'''        RSTTRXFILE!line_no = p.Item(I).Item("LN_NO")
'''        RSTTRXFILE!ITEM_CODE = p.Item(I).Item("ITEM_CODE")
'''        RSTTRXFILE!ITEM_NAME = p.Item(I).Item("ITEM_NAME")
'''        RSTTRXFILE!item_uprice = p.Item(I).Item("U_PRICE")
'''        RSTTRXFILE!ITEM_QTY = p.Item(I).Item("QTY")
'''        RSTTRXFILE!item_tprice = p.Item(I).Item("T_PRICE")
'''        RSTTRXFILE!c_date = Format(Date, "dd/mm/yyyy")
'''        RSTTRXFILE.Update
'''    Next I
'''    RSTTRXFILE.Close
'''    Set RSTTRXFILE = Nothing
'''
'''    db.Execute "delete from orders where COMP_CODE <> '" & Trim(customercode) & "' "
'''    db.BeginTrans
'''    Set RSTordtrxfile = New ADODB.Recordset
'''    RSTordtrxfile.Open "Select distinct ord_no from orders", db, adOpenStatic, adLockReadOnly
'''    'RSTordtrxfile.Open "SELECT DISTINCT VCH_NO, TRX_TYPE, TRX_YEAR From TEMPCN WHERE VCH_DATE <= '" & Format(DTTO.Value, "yyyy/mm/dd") & "' AND VCH_DATE >= '" & Format(DTFROM.Value, "yyyy/mm/dd") & "' ORDER BY TRX_TYPE,VCH_DATE,VCH_NO", db, adOpenStatic, adLockReadOnly
'''    Do Until RSTordtrxfile.EOF
'''
'''        Set RSTordtrxfile1 = New ADODB.Recordset
'''        RSTordtrxfile1.Open "Select * from orders where ord_no = '" & RSTordtrxfile!ord_no & "' ", db, adOpenStatic, adLockReadOnly
'''
'''        Set RSTTRXFILE = New ADODB.Recordset
'''        RSTTRXFILE.Open "Select * From ord_mast WHERE ord_no= (SELECT MAX(ord_no) FROM ord_mast)", db, adOpenStatic, adLockOptimistic, adCmdText
'''        If (RSTTRXFILE.EOF And RSTTRXFILE.BOF) Then
'''            BILL_NO = 1
'''        Else
'''            BILL_NO = RSTTRXFILE!ord_no + 1
'''        End If
'''
'''        RSTTRXFILE.AddNew
'''        RSTTRXFILE!ord_no = BILL_NO
'''        RSTTRXFILE!ACT_CODE = RSTordtrxfile1!ACT_CODE
'''        RSTTRXFILE!ACT_NAME = RSTordtrxfile1!ACT_NAME
'''        RSTTRXFILE!act_address = ""
'''        RSTTRXFILE!act_phone = ""
'''        RSTTRXFILE!C_USER_ID = RSTordtrxfile1!USER_ID
'''        RSTTRXFILE!C_USER_DATE = Format(Date, "DD/MM/YYYY")
'''        RSTTRXFILE.Update
'''        RSTTRXFILE.Close
'''        Set RSTTRXFILE = Nothing
'''
'''        RSTordtrxfile1.Close
'''        Set RSTordtrxfile1 = Nothing
'''
'''        db.Execute "delete from  ord_trxfile where ord_no = '" & BILL_NO & "'"
'''
'''        Set RSTTRXFILE = New ADODB.Recordset
'''        RSTTRXFILE.Open "Select * from ord_trxfile", db, adOpenStatic, adLockOptimistic, adCmdText
'''        Set RSTordtrxfile1 = New ADODB.Recordset
'''        RSTordtrxfile1.Open "Select * from orders where ord_no = '" & RSTordtrxfile!ord_no & "' ", db, adOpenStatic, adLockReadOnly
'''        Do Until RSTordtrxfile1.EOF
'''            RSTTRXFILE.AddNew
'''            RSTTRXFILE!ord_no = BILL_NO
'''            RSTTRXFILE!line_no = RSTordtrxfile1!line_no
'''            RSTTRXFILE!ITEM_CODE = RSTordtrxfile1!ITEM_CODE
'''            RSTTRXFILE!ITEM_NAME = RSTordtrxfile1!ITEM_NAME
'''            RSTTRXFILE!ITEM_QTY = RSTordtrxfile1!ITEM_QTY
'''            'RSTTRXFILE!item_uprice = RSTordtrxfile1!item_uprice
'''            'RSTTRXFILE!item_tprice = RSTordtrxfile1!item_tprice
'''            RSTTRXFILE.Update
'''            RSTordtrxfile1.MoveNext
'''        Loop
'''        RSTTRXFILE.Close
'''        Set RSTTRXFILE = Nothing
'''
'''        RSTordtrxfile1.Close
'''        Set RSTordtrxfile1 = Nothing
'''
'''        Set http = CreateObject("WinHttp.WinHttprequest.5.1")
'''        URL = "https://www.ezbiz.co.in/ezbiz/delorder.php?order_no=" & RSTordtrxfile!ord_no
'''        http.Open "Get", URL, False
'''        http.send
'''
'''        RSTordtrxfile.MoveNext
'''    Loop
'''    db.CommitTrans
'''
'''    Screen.MousePointer = vbNormal
'''    If MsgBox("Do you want to update customer data?", vbYesNo + vbDefaultButton2, "EzBiz.....") = vbYes Then
'''        Screen.MousePointer = vbHourglass
'''        Set http = CreateObject("WinHttp.WinHttprequest.5.1")
'''        Set RSTTRXFILE = New ADODB.Recordset
'''        RSTTRXFILE.Open "SELECT * FROM CUSTMAST ", db, adOpenStatic, adLockReadOnly, adCmdText
'''        Do Until RSTTRXFILE.EOF
'''            If RSTTRXFILE!Type = "W" Then
'''                CUSTTYPE = 2
'''            Else
'''                CUSTTYPE = 1
'''            End If
'''
'''            URL = "https://ezbiz.co.in/ezbiz/addcust.php?comp_code=" & Trim(customercode) & "&act_code=" & RSTTRXFILE!ACT_CODE & "&act_name=" & RSTTRXFILE!ACT_NAME & "&address=" & RSTTRXFILE!Address & "&telno=null&area=null&contact_person=null&pymt_period=1&ytd_cr=null&cust_type=" & CUSTTYPE & "&create_date=null&c_user_id=null&modify_date=null&m_user_id=null&cust_igst=null&agent_code=null&agent_code=null&agent_name=null&pymt_limit=null"
'''            http.Open "Get", URL, False
'''            http.send
'''
'''            RSTTRXFILE.MoveNext
'''        Loop
'''        RSTTRXFILE.Close
'''        Set RSTTRXFILE = Nothing
'''    End If
'''
'''    Screen.MousePointer = vbNormal
'''    If MsgBox("Do you want to update item list? This may take some time to complete.", vbYesNo + vbDefaultButton2, "EzBiz.....") = vbYes Then
'''        Screen.MousePointer = vbHourglass
'''        Set http = CreateObject("WinHttp.WinHttprequest.5.1")
'''        Set RSTTRXFILE = New ADODB.Recordset
'''        RSTTRXFILE.Open "SELECT * FROM ITEMMAST WHERE (ISNULL(DEAD_STOCK) OR DEAD_STOCK ='N')", db, adOpenStatic, adLockReadOnly, adCmdText
'''        Do Until RSTTRXFILE.EOF
'''            URL = "https://ezbiz.co.in/ezbiz/additem.php?comp_code=" & Trim(customercode) & _
'''            "&item_code=" & RSTTRXFILE!ITEM_CODE & _
'''            "&item_name=" & RSTTRXFILE!ITEM_NAME & _
'''            "&category=" & RSTTRXFILE!Category & _
'''            "&item_cost=" & RSTTRXFILE!ITEM_COST & _
'''            "&mrp=" & RSTTRXFILE!MRP & _
'''            "&sales_tax=" & RSTTRXFILE!SALES_TAX & _
'''            "&ptr=" & RSTTRXFILE!PTR & _
'''            "&close_qty=" & RSTTRXFILE!CLOSE_QTY & _
'''            "&close_val=" & RSTTRXFILE!CLOSE_VAL & _
'''            "&manufacturer=" & RSTTRXFILE!MANUFACTURER & _
'''            "&create_date=" & Format(Date, "DD/MM/YYYY") & _
'''            "&c_user_id=" & frmLogin.rs!USER_ID & _
'''            "&price1=" & IIf(IsNull(RSTTRXFILE!P_RETAIL), 0, RSTTRXFILE!P_RETAIL) & _
'''            "&price2=" & IIf(IsNull(RSTTRXFILE!P_WS), 0, RSTTRXFILE!P_WS) & _
'''            "&price3=" & IIf(IsNull(RSTTRXFILE!P_VAN), 0, RSTTRXFILE!P_VAN) & _
'''            "&price4=" & IIf(IsNull(RSTTRXFILE!PRICE5), 0, RSTTRXFILE!PRICE5) & _
'''            "&price5=" & IIf(IsNull(RSTTRXFILE!PRICE6), 0, RSTTRXFILE!PRICE6) & _
'''            "&price6=" & IIf(IsNull(RSTTRXFILE!PRICE7), 0, RSTTRXFILE!PRICE7) & _
'''            "&pack_type=" & RSTTRXFILE!PACK_TYPE & _
'''            "&cust_disc=" & IIf(IsNull(RSTTRXFILE!CUST_DISC), 0, RSTTRXFILE!CUST_DISC) & _
'''            "&un_bill=" & RSTTRXFILE!UN_BILL & _
'''            "&item_net_cost=" & IIf(IsNull(RSTTRXFILE!ITEM_NET_COST), 0, RSTTRXFILE!ITEM_NET_COST) & _
'''            "&barcode=" & IIf(IsNull(RSTTRXFILE!BARCODE), "", RSTTRXFILE!BARCODE)
'''
'''            '"&pack_type=null&cust_disc=0&un_bill=0&item_net_cost=0&barcode=null"
'''            http.Open "Get", URL, False
'''            http.send
'''
'''            RSTTRXFILE.MoveNext
'''        Loop
'''        RSTTRXFILE.Close
'''        Set RSTTRXFILE = Nothing
'''    End If
''''    For i = 1 To GrdOrder.rows - 1
''''        Set RSTordtrxfile = New ADODB.Recordset
''''        RSTordtrxfile.Open "Select * FROM ord_trxfile WHERE ord_no = " & Val(txtBillNo.text) & " AND line_no = " & Val(GrdOrder.TextMatrix(i, 0)) & "", db, adOpenStatic, adLockOptimistic, adCmdText
''''        If (RSTordtrxfile.EOF And RSTordtrxfile.BOF) Then
''''            RSTordtrxfile.AddNew
''''            RSTordtrxfile!ord_no = txtBillNo.text
''''            RSTordtrxfile!line_no = Val(GrdOrder.TextMatrix(i, 0))
''''    '        RSTordtrxfile!C_USER_ID = frmLogin.rs!USER_ID
''''    '        RSTordtrxfile!CREATE_DATE = Format(Date, "DD/MM/YYYY")F
''''        End If
''''        RSTordtrxfile!item_code = GrdOrder.TextMatrix(i, 3)
''''        RSTordtrxfile!item_name = GrdOrder.TextMatrix(i, 1)
''''        RSTordtrxfile!item_qty = Val(GrdOrder.TextMatrix(i, 2))
''''        RSTordtrxfile.Update
''''    Next i
''''    db.CommitTrans
'''
''''ErrHand:
''''    MsgBox Err.Description, , "EzBiz"
'''    'MsgBox http.responseText
    
    Dim cmd As String

'    'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " > " & App.Path & "\Backup\" & strBackupEXT
'    'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -hlocalhost -uroot -p###%%database%%###ret invsoft transmast" & "|" & App.Path & "\mysql.exe" & Chr(34) & " -hlocalhost -uroot -p###%%database%%###ret tempdb transmast001"
'
'
'    'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " transmast | sed -e 's/ `transmast` / `transmastnew` > " & App.Path & "\Backup\dump.txt"
'    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost --routines --comments " & dbase1 & " transmast > " & App.Path & "\Backup\dump.txt"
'
'    'mysqldump -u root -p database_name \ | mysql -h other-host.com database_name
'    Call execCommand(cmd)
'
'    DoEvents
'    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost invdb < " & App.Path & "\Backup\dump.txt"
'    Call execCommand(cmd)
'
'    db.Execute "Alter table transmast Rename to transmast123"
'    Exit Function
    
    Dim dbexport As New ADODB.Connection
    Dim strConn As String
    'Dim servername As String
    Dim TABLE_NAME As String

    If Not FileExists(App.Path & "\mysqldump.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Function
    End If

    If Not FileExists(App.Path & "\mysql.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Function
    End If

    Screen.MousePointer = vbHourglass
    'servername = "162.241.123.157"
    'servername = "www.edocbew.com"
    'servername = "www.gold.primecrown.net"
    On Error GoTo ERRHAND
    'Exit Function
    db.Execute "DROP TABLE if exists CUSTMASTTEMP "
    db.Execute "create table CUSTMASTTEMP (COMP_CODE varchar(10) NULL, ACT_CODE varchar(15) NOT NULL, ACT_NAME varchar(150) NULL, ADDRESS varchar(150) NULL, TELNO varchar(15) NULL, AREA varchar(200) NULL, CONTACT_PERSON varchar(35) NULL, PYMT_PERIOD int(11) DEFAULT 0, YTD_CR double NULL, CUST_TYPE varchar(1) NULL, CREATE_DATE varchar(15) NULL, C_USER_ID varchar(8) NULL, MODIFY_DATE varchar(15) NULL, M_USER_ID varchar(8) NULL, CUST_IGST varchar(1) NULL, AGENT_CODE varchar(6) NULL, AGENT_NAME varchar(35) NULL, PYMT_LIMIT int(8) NULL, PRIMARY KEY (ACT_CODE)) ENGINE = MyISAM"
    db.Execute "INSERT INTO `CUSTMASTTEMP` SELECT REMARKS, ACT_CODE, ACT_NAME, ADDRESS, TELNO, AREA, CONTACT_PERSON, PYMT_PERIOD, YTD_CR, TYPE, CREATE_DATE, C_USER_ID, MODIFY_DATE, M_USER_ID, CUST_IGST, AGENT_CODE, AGENT_NAME, PYMT_LIMIT  FROM `CUSTMAST`"
    db.Execute "Update CUSTMASTTEMP SET COMP_CODE = '" & customercode & "' "
    db.Execute "Update CUSTMASTTEMP SET CUST_TYPE = '1' where CUST_TYPE = 'R'"
    db.Execute "Update CUSTMASTTEMP SET CUST_TYPE = '2' where CUST_TYPE = 'W'"
    db.Execute "Update CUSTMASTTEMP SET CUST_TYPE = '3' where CUST_TYPE = 'V'"
    db.Execute "Update CUSTMASTTEMP SET CUST_TYPE = '4' where CUST_TYPE = 'M'"
    
    
    
    'strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & servername & ";Port=3306;Database=ezbizco_data;User=ezbizco_admin; Password=ezbizadmin@2022; Option=2;"
    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & serveraddress & ";Port=3306;Database=ezbizkay_dataV2;User=ezbizkay_adminV2; Password=Ezbizadmin@12345; Option=2;"
    dbexport.Open strConn
    dbexport.CursorLocation = adUseClient
    
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " CUSTMASTTEMP > " & App.Path & "\Backup\dump.txt"
    'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " CUSTMASTTEMP | Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uezbizkay_adminV2 -pEzbizadmin@12345 -h162.241.123.157 ezbizkay_dataV2
'    cmd = App.Path & "\mysqldump.exe" & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " CUSTMASTTEMP | mysql -h 162.241.123.157 -u ezbizkay_adminV2 -pEzbizadmin@12345 ezbizkay_dataV2"
    Call execCommand(cmd)

    'dbexport.Execute "TRUNCATE TABLE custmast where COMP_CODE = '" & customercode & "'"
    
    DoEvents
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -pEzbizadmin@12345 -h " & serveraddress & " 162.241.123.157 < " & App.Path & "\Backup\dump.txt"
    
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uezbizkay_adminV2 -pEzbizadmin@12345 -h162.241.123.157 ezbizkay_dataV2 <" & App.Path & "\Backup\dump.txt"
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost invdb < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)
    
    dbexport.Execute "DELETE FROM custmast where COMP_CODE = '" & customercode & "'"
    dbexport.Execute "INSERT INTO `custmast` SELECT COMP_CODE, ACT_CODE, ACT_NAME, ADDRESS, TELNO, AREA, CONTACT_PERSON, PYMT_PERIOD, YTD_CR, CUST_TYPE, CREATE_DATE, C_USER_ID, MODIFY_DATE, M_USER_ID, CUST_IGST, AGENT_CODE, AGENT_NAME, PYMT_LIMIT  FROM `custmasttemp`"
    
'
'    '1. transmast
''    TABLE_NAME = "transmast" & customercode
'    dbexport.Execute "TRUNCATE TABLE `custmast` "
'    'dbexport.Execute "CREATE TABLE " & TABLE_NAME & " LIKE db.transmast "
'    dbexport.Execute "INSERT INTO `custmast` SELECT ACT_CODE, ACT_NAME, ADDRESS, AGENT_CODE, AGENT_NAME, AREA  FROM db.`transmast`"
''    USE db2;
''
''CREATE TABLE table2 LIKE db1.table1;
''
''INSERT INTO table2
''    SELECT * FROM db1.table1;
    
    
    
    '2 - itemmast
    db.Execute "DROP TABLE if exists ITEMMASTTEMP "
    db.Execute "create table ITEMMASTTEMP (COMP_CODE varchar(10) NULL, ITEM_CODE varchar(20) NOT NULL, ITEM_NAME varchar(200) NULL, CATEGORY varchar(50) NULL, ITEM_COST double NULL, MRP double NULL, SALES_TAX double NULL, PTR double NULL, CLOSE_QTY double NULL, CLOSE_VAL double NULL, MANUFACTURER varchar(25) NULL, CREATE_DATE varchar(15) NULL, C_USER_ID varchar(8) NULL, MODIFY_DATE varchar(15) NULL, M_USER_ID varchar(8) NULL, PRICE1 double NULL, PRICE2 double NULL, PRICE3 double NULL, PRICE4 double NULL, PRICE5 double NULL, PRICE6 double NULL, PACK_TYPE varchar(50) NULL, CUST_DISC double NULL, UN_BILL varchar(1) NULL, ITEM_NET_COST double NULL, BARCODE varchar(30) NULL, PRIMARY KEY (ITEM_CODE)) ENGINE = MyISAM"
    
    db.Execute "INSERT INTO `ITEMMASTTEMP` SELECT REMARKS, ITEM_CODE, ITEM_NAME, CATEGORY, ITEM_COST, MRP, SALES_TAX, PTR, CLOSE_QTY, CLOSE_VAL, MANUFACTURER, CREATE_DATE, C_USER_ID, MODIFY_DATE, M_USER_ID, P_RETAIL, P_WS, P_VAN, PRICE5, PRICE6, PRICE7, PACK_TYPE, CUST_DISC, UN_BILL, ITEM_NET_COST, BARCODE FROM `ITEMMAST`"
            
    db.Execute "Update ITEMMASTTEMP  SET PRICE1 = 0 where isnull(PRICE1)"
    db.Execute "Update ITEMMASTTEMP  SET PRICE2 = 0 where isnull(PRICE2)"
    db.Execute "Update ITEMMASTTEMP  SET PRICE3 = 0 where isnull(PRICE3)"
    db.Execute "Update ITEMMASTTEMP  SET PRICE4 = 0 where isnull(PRICE4)"
    db.Execute "Update ITEMMASTTEMP  SET PRICE5 = 0 where isnull(PRICE5)"
    db.Execute "Update ITEMMASTTEMP  SET PRICE6 = 0 where isnull(PRICE6)"
    db.Execute "Update ITEMMASTTEMP  SET MRP = 0 where isnull(MRP)"
    
    db.Execute "Update ITEMMASTTEMP  SET COMP_CODE = '" & customercode & "' "
    'strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & serveraddress & ";Port=3306;Database=ezbizco_data;User=ezbizco_admin; Password=ezbizadmin@2022; Option=2;"
'    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & serveraddress & ";Port=3306;Database=ezbizkay_dataV2;User=ezbizkay_adminV2; Password=Ezbizadmin@12345; Option=2;"
'    dbexport.Open strConn
'    dbexport.CursorLocation = adUseClient
    
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " ITEMMASTTEMP > " & App.Path & "\Backup\dump.txt"
    'cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " CUSTMASTTEMP | Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uezbizkay_adminV2 -pEzbizadmin@12345 -h162.241.123.157 ezbizkay_dataV2
'    cmd = App.Path & "\mysqldump.exe" & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " CUSTMASTTEMP | mysql -h 162.241.123.157 -u ezbizkay_adminV2 -pEzbizadmin@12345 ezbizkay_dataV2"
    Call execCommand(cmd)

    'dbexport.Execute "TRUNCATE TABLE custmast where COMP_CODE = '" & customercode & "'"
    
    DoEvents
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -pEzbizadmin@12345 -h " & serveraddress & " 162.241.123.157 < " & App.Path & "\Backup\dump.txt"
    
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uezbizkay_adminV2 -pEzbizadmin@12345 -h162.241.123.157 ezbizkay_dataV2 <" & App.Path & "\Backup\dump.txt"
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost invdb < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)
    
    dbexport.Execute "DELETE FROM itemmast where COMP_CODE = '" & customercode & "'"
    dbexport.Execute "INSERT INTO `itemmast` SELECT COMP_CODE, ITEM_CODE, ITEM_NAME, CATEGORY, ITEM_COST, MRP, SALES_TAX, PTR, CLOSE_QTY, CLOSE_VAL, MANUFACTURER, CREATE_DATE, C_USER_ID, MODIFY_DATE, M_USER_ID, PRICE1, PRICE2, PRICE3, PRICE4, PRICE5, PRICE6, PACK_TYPE, CUST_DISC, UN_BILL, ITEM_NET_COST, BARCODE FROM `itemmasttemp`"
    
    
    
    dbexport.Close
    Set dbexport = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Synced successfully", , "EzBiz"
    Exit Function
 '   db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    
    
    
    
    '1. transmast
    TABLE_NAME = "transmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `transmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `transmast`"

    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -pEzbizadmin@12345 -h " & serveraddress & " 162.241.123.157 < " & App.Path & "\Backup\dump.txt"
    
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uezbizkay_adminV2 -pEzbizadmin@12345 -h162.241.123.157 ezbizkay_dataV2 <" & App.Path & "\Backup\dump.txt"
    'cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -hlocalhost invdb < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '2. ASTRXMAST
    TABLE_NAME = "astrxmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `astrxmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `astrxmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '3. TRXEXPMAST
    TABLE_NAME = "trxexpmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxexpmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxexpmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & ""
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '4. trxexp_mast
    TABLE_NAME = "trxexp_mast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxexp_mast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxexp_mast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '5. trxincmast
    TABLE_NAME = "trxincmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxincmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxincmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '6. trxmast
    TABLE_NAME = "trxmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '7. returnmast
    TABLE_NAME = "returnmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `returnmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `returnmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '8. rtrxfile
    TABLE_NAME = "rtrxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `rtrxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `rtrxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '9. crdtpymt
    TABLE_NAME = "crdtpymt" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `crdtpymt` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `crdtpymt`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '10. dbtpymt
    TABLE_NAME = "dbtpymt" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `dbtpymt` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `dbtpymt`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '11. actmast
    TABLE_NAME = "actmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `actmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `actmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '12. cashatrxfile
    TABLE_NAME = "cashatrxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `cashatrxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `cashatrxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '13. bank_trx
    TABLE_NAME = "bank_trx" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `bank_trx` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `bank_trx`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '14. trxfile
    TABLE_NAME = "trxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u ezbizkay_adminV2 -Ezbizadmin@12345 -h " & serveraddress & " ezbizkay_adminV2_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)
'''
'''    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    Dim rstTRXMAST  As ADODB.Recordset
'    Dim RSTTRXFILE  As ADODB.Recordset
'
'    TABLE_NAME = "transmast" & customercode
'    dbprint.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'     dbprint.Execute "CREATE TABLE dbexport.TEST SELECT * FROM DB.transmast"
'
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `NET_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    Set rstTRXMAST = New ADODB.Recordset
'    rstTRXMAST.Open "Select * From transmast", db, adOpenForwardOnly
'
'    Set RSTTRXFILE = New ADODB.Recordset
'    RSTTRXFILE.Open "Select * From " & TABLE_NAME & "", dbexport, adOpenStatic, adLockOptimistic, adCmdText
'    Do Until rstTRXMAST.EOF
'        dbexport.Execute "Update " & TABLE_NAME & " set TRX_TYPE = '" & rstTRXMAST!TRX_TYPE & "', VCH_NO = " & rstTRXMAST!VCH_NO & ", TRX_YEAR = '" & rstTRXMAST!TRX_YEAR & "', NET_AMOUNT = " & rstTRXMAST!NET_AMOUNT & "  WHERE TRX_TYPE = '" & rstTRXMAST!TRX_TYPE & "' AND VCH_NO = " & rstTRXMAST!VCH_NO & " AND TRX_YEAR = '" & rstTRXMAST!TRX_YEAR & "' "
'        'REF_NO='" & Left(INVDETAILS, 20) & "'
''        rstTRXFILE.AddNew
''        rstTRXFILE!TRX_TYPE = rstTRXMAST!TRX_TYPE
''        rstTRXFILE!VCH_NO = rstTRXMAST!VCH_NO
''        rstTRXFILE!TRX_YEAR = rstTRXMAST!TRX_YEAR
''        rstTRXFILE!VCH_DATE = rstTRXMAST!VCH_DATE
''        rstTRXFILE!NET_AMOUNT = rstTRXMAST!NET_AMOUNT
''
''        rstTRXFILE.Update
'        rstTRXMAST.MoveNext
'    Loop
'    RSTTRXFILE.Close
'    Set RSTTRXFILE = Nothing
'
'    rstTRXMAST.Close
'    Set rstTRXMAST = Nothing
'
'
'    TABLE_NAME = "ASTRXMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `NET_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "TRXEXPMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `VCH_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "TRXEXP_MAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `VCH_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "TRXINCMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `VCH_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "TRXMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_DATE` date NULL , `VCH_AMOUNT` DOUBLE NULL , `NET_AMOUNT` DOUBLE NULL , `DISC_PERS` DOUBLE NULL, `SLSM_CODE` VARCHAR(4) NULL, `COMM_AMT` DOUBLE NULL, `DISCOUNT` DOUBLE NULL, `PAY_AMOUNT` DOUBLE NULL, `HANDLE` DOUBLE NULL, `FRIEGHT` DOUBLE NULL, `ACT_CODE` VARCHAR(30) NULL  ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "RETURNMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`VCH_AMOUNT` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "RTRXFILE" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`TRX_TOTAL` DOUBLE NOT NULL ,PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
'
'     TABLE_NAME = "CRDTPYMT" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`CR_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`ACT_NAME` varchar(150) NULL ,`RCPT_AMOUNT` DOUBLE NULL ,`RCPT_DATE` datetime NULL ,`INV_NO` DOUBLE NULL ,`INV_DATE` date NULL ,`REF_NO` varchar(30) NULL ,PRIMARY KEY ( `CR_NO` , `TRX_TYPE`)) ENGINE = MyISAM;"
'
'    TABLE_NAME = "DBTPYMT" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`CR_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`ACT_NAME` varchar(150) NULL ,`RCPT_AMT` DOUBLE NULL ,`RCPT_DATE` datetime NULL ,`INV_NO` DOUBLE NULL ,`INV_DATE` date NULL ,`REF_NO` varchar(100) NULL ,PRIMARY KEY ( `CR_NO` , `TRX_TYPE` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "ACTMAST" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`ACT_CODE` VARCHAR( 6 ) NOT NULL ,`OPEN_DB` DOUBLE NULL ,PRIMARY KEY ( `ACT_CODE` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "CASHATRXFILE" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`REC_NO` DOUBLE NOT NULL ,`INV_TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`INV_TYPE` varchar(2) NULL ,`INV_NO` INT (11) NOT NULL ,`VCH_DATE` date NULL ,`CHECK_FLAG` varchar(1) NULL ,`AMOUNT` double NULL , PRIMARY KEY ( `TRX_TYPE` , `REC_NO`, `INV_TRX_TYPE`, `INV_TYPE`, `INV_NO` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "BANK_TRX" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL , `BANK_CODE` VARCHAR( 4 ) NOT NULL ,`TRX_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL , `BILL_TRX_TYPE` VARCHAR( 2 ) NOT NULL , `BNK_SL_NO` VARCHAR( 4 ) NOT NULL ,`TRX_DATE` date NULL,`TRX_AMOUNT` DOUBLE NULL,`BANK_NAME` VARCHAR(35) NULL,`ACT_NAME` VARCHAR(150) NULL ,PRIMARY KEY (`BANK_CODE`, `TRX_NO`, `TRX_TYPE`, `BILL_TRX_TYPE`, `TRX_YEAR` )) ENGINE = MyISAM;"
'
'    TABLE_NAME = "TRXFILE" & customercode
'    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
'    dbexport.Execute "CREATE TABLE " & TABLE_NAME & " (`TRX_TYPE` VARCHAR( 2 ) NOT NULL ,`VCH_NO` DOUBLE NOT NULL, `LINE_NO` DOUBLE NOT NULL ,`TRX_YEAR` VARCHAR( 4 ) NOT NULL ,`TRX_TOTAL` DOUBLE NULL , `PTR` DOUBLE NULL , `LINE_DISC` DOUBLE NULL, `KFC_TAX` DOUBLE NULL, `SALES_TAX` DOUBLE NULL, `CESS_PER` DOUBLE NULL, `CESS_AMT` DOUBLE NULL, `ITEM_COST` DOUBLE NULL, `QTY` DOUBLE NULL, `PUR_TAX` DOUBLE NULL, `UN_BILL` VARCHAR(1) NULL, `VCH_DATE` date NULL, PRIMARY KEY ( `TRX_TYPE` , `VCH_NO` , `LINE_NO` , `TRX_YEAR` )) ENGINE = MyISAM;"
    
    dbexport.Close
    Set dbexport = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Synced successfully", , "EzBiz"
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    MsgBox err.Description, , "EzBiz"
'    On Error Resume Next
'    db.RollbackTrans
End Function

Public Sub execCommand(ByVal cmd As String)
    Dim result  As Long
    Dim lPid    As Long
    Dim lHnd    As Long
    Dim lRet    As Long
 
    cmd = "cmd /c " & cmd
    result = Shell(cmd, vbHide)
 
    lPid = result
    If lPid <> 0 Then
        lHnd = OpenProcess(SYNCHRONIZE, 0, lPid)
        If lHnd <> 0 Then
            lRet = WaitForSingleObject(lHnd, INFINITE)
            CloseHandle (lHnd)
        End If
    End If
End Sub

Public Function errcodes(errc As Integer)
    Select Case errc
        Case 1
            MsgBox "ER_JSON_TABLE_SCALAR_EXPECTED: Can't store an array or an object in the scalar column '%s' of JSON_TABLE '%s'.", vbCritical, "EzBiz"
        Case 2
            MsgBox "WARN_VERS_PART_FULL: Versioned table %`s.%`s: last HISTORY partition (%`s) is out of %s, need more HISTORY partitions", vbCritical, "EzBiz"
        Case 3
            MsgBox "ER_OUT_OF_RESOURCES: Out of memory; check if mysqld or some other process uses all available memory; if not, you may have to use 'ulimit' to allow mysqld to use more memory or you can add more swap space", vbCritical, "EzBiz"
        Case 4
            MsgBox "ER_WRONG_AUTO_KEY: Incorrect table definition; there can be only one auto column and it must be defined as a key", vbCritical, "EzBiz"
        Case 5
            MsgBox "ER_TOO_BIG_ROWSIZE: Row size too large. The maximum row size for the used table type, not counting BLOBs, is %ld. You have to change some columns to TEXT or BLOBs", vbCritical, "EzBiz"
        Case 6
            MsgBox "ER_NULL_COLUMN_IN_INDEX: Table handler doesn't support NULL in given index. Please change column '%s' to be NOT NULL or use another handler", vbCritical, "EzBiz"
        Case 7
            MsgBox "Error: 1153 SQLSTATE: 08S01 (ER_NET_PACKET_TOO_LARGE): Got a packet bigger than 'max_allowed_packet' bytes", vbCritical, "EzBiz"
        Case 8
            MsgBox "Error: 1168 SQLSTATE: HY000 (ER_WRONG_MRG_TABLE): Unable to open underlying table which is differently defined or of non-MyISAM type or doesn't exist ", vbCritical, "EzBiz"
        Case 9
            MsgBox "Error: 1192 SQLSTATE: HY000 (ER_LOCK_OR_ACTIVE_TRANSACTION): Can't execute the given command because you have active locked tables or an active transaction ", vbCritical, "EzBiz"
        Case 10
            MsgBox "Error: 1199 SQLSTATE: HY000 (ER_SLAVE_NOT_RUNNING): This operation requires a running slave; configure slave and do START SLAVE ", vbCritical, "EzBiz"
        Case 11
            MsgBox "Error: 1200 SQLSTATE: HY000 (ER_BAD_SLAVE): The server is not configured as slave; fix in config file or with CHANGE MASTER TO", vbCritical, "EzBiz"
        Case 12
            MsgBox "Error: 1206 SQLSTATE: HY000 (ER_LOCK_TABLE_FULL): The total number of locks exceeds the lock table size", vbCritical, "EzBiz"
        Case 13
            MsgBox "Error: 1207 SQLSTATE: 25000 (ER_READ_ONLY_TRANSACTION): Update locks cannot be acquired during a READ UNCOMMITTED transaction", vbCritical, "EzBiz"
        Case 14
            MsgBox "mysql> SET GLOBAL innodb_version='My InnoDB Version': ERROR 1238 (HY000): Variable 'innodb_version' is a read only variable", vbCritical, "EzBiz"
        Case 15
            MsgBox "Error: 1258 SQLSTATE: HY000 (ER_ZLIB_Z_BUF_ERROR): ZLIB: Not enough room in the output buffer (probably, length of uncompressed data was corrupted) ", vbCritical, "EzBiz"
        Case 16
            MsgBox "Error: 1274 SQLSTATE: HY000 (ER_SLAVE_IGNORED_SSL_PARAMS): SSL parameters in CHANGE MASTER are ignored because this MySQL slave was compiled without SSL support; they can be used later if MySQL slave with SSL is started ", vbCritical, "EzBiz"
        Case 17
            MsgBox "Error: 1278 SQLSTATE: HY000 (ER_MISSING_SKIP_SLAVE): It is recommended to use --skip-slave-start when doing step-by-step replication with START SLAVE UNTIL; otherwise, you will get problems if you get an unexpected slave's mysqld restart ", vbCritical, "EzBiz"
        Case 18
            MsgBox "Error: 1293 SQLSTATE: HY000 (ER_TOO_MUCH_AUTO_TIMESTAMP_COLS): Incorrect table definition; there can be only one TIMESTAMP column with CURRENT_TIMESTAMP in DEFAULT or ON UPDATE clause ", vbCritical, "EzBiz"
        Case 19
            MsgBox "Error: 1315 SQLSTATE: 42000 (ER_UPDATE_LOG_DEPRECATED_IGNORED): The update log is deprecated and replaced by the binary log; SET SQL_LOG_UPDATE has been ignored. This option will be removed in MySQL 5.6. ", vbCritical, "EzBiz"
        Case 20
            MsgBox "Error: 1345 SQLSTATE: HY000 (ER_VIEW_NO_EXPLAIN): EXPLAIN/SHOW can not be issued; lacking privileges for underlying table ", vbCritical, "EzBiz"
        Case 21
            MsgBox "Error: 1356 SQLSTATE: HY000 (ER_VIEW_INVALID): View '%s.%s' references invalid table(s) or column(s) or function(s) or definer/invoker of view lack rights to use them ", vbCritical, "EzBiz"
        Case 22
            MsgBox "Error: 1399 SQLSTATE: XAE07 (ER_XAER_RMFAIL): XAER_RMFAIL: The command cannot be executed when global transaction is in the %s state ", vbCritical, "EzBiz"
        Case 23
            MsgBox "Error: 1417 SQLSTATE: HY000 (ER_FAILED_ROUTINE_BREAK_BINLOG): A routine failed and has neither NO SQL nor READS SQL DATA in its declaration and binary logging is enabled; if non-transactional tables were updated, the binary log will miss their changes ", vbCritical, "EzBiz"
        Case 24
            MsgBox "[ERROR_EXCL_SEM_ALREADY_OWNED (0x65)] The exclusive semaphore is owned by another process.", vbCritical, "EzBiz"
        Case 25
            MsgBox "ERROR_ARENA_TRASHED: The storage control blocks were destroyed", vbCritical, "EzBiz"
        Case 26
            MsgBox "EError: 1420 SQLSTATE: HY000 (ER_EXEC_STMT_WITH_OPEN_CURSOR): You can 't execute a prepared statement which has an open cursor associated with it. Reset the statement to re-execute it. ", vbCritical, "EzBiz"
        Case 27
            MsgBox "Error: 1418 SQLSTATE: HY000 (ER_BINLOG_UNSAFE_ROUTINE): This function has none of DETERMINISTIC, NO SQL, or READS SQL DATA in its declaration and binary logging is enabled (you *might* want to use the less safe log_bin_trust_function_creators variable) ", vbCritical, "EzBiz"
        Case 28
            MsgBox "Error: 1442 SQLSTATE: HY000 (ER_CANT_UPDATE_USED_TABLE_IN_SF_OR_TRG): Can 't update table '%s' in stored function/trigger because it is already used by statement which invoked this stored function/trigger. ", vbCritical, "EzBiz"
        Case 29
            MsgBox "Error: 1447 SQLSTATE: HY000 (ER_VIEW_FRM_NO_USER): View '%s'.'%s' has no definer information (old table format). Current user is used as definer. Please recreate the view! ", vbCritical, "EzBiz"
        Case 30
            MsgBox "Error: 1475 SQLSTATE: HY000 (ER_AMBIGUOUS_FIELD_TERM): First character of the FIELDS TERMINATED string is ambiguous; please use non-optional and non-empty FIELDS ENCLOSED BY ", vbCritical, "EzBiz"
        Case 31
            MsgBox "Error: 1486 SQLSTATE: HY000 (ER_WRONG_EXPR_IN_PARTITION_FUNC_ERROR): Constant, random or timezone-dependent expressions in (sub)partitioning function are not allowed ", vbCritical, "EzBiz"
        Case Else
            MsgBox "ERROR_SECTOR_NOT_FOUND: The drive cannot find the sector requested", vbCritical, "EzBiz"
    End Select
End Function

Public Function IsConnected() As Boolean

    On Error GoTo err
    IsConnected = InternetGetConnectedState(0&, 0&)

Exit Function

err:
    IsConnected = True

End Function

Public Function export_db2()
    
    Dim dbexport As New ADODB.Connection
    Dim strConn As String
    Dim dbserver As String
    Dim TABLE_NAME As String
    Dim cmd As String
    
    If customercode = "" Then Exit Function
    dbserver = "inv" & customercode
    If Not FileExists(App.Path & "\mysqldump.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Function
    End If

    If Not FileExists(App.Path & "\mysql.exe") Then
        MsgBox "File not exists", , "EzBiz"
        Exit Function
    End If

    Screen.MousePointer = vbHourglass
    On Error GoTo ERRHAND
        
    strConn = "Driver={MySQL ODBC 5.1 Driver};Server=" & serveraddress & ";Port=3306;Database=tempdb;User=root; Password=###%%database%%###ret; Option=2;"
    dbexport.Open strConn
    dbexport.CursorLocation = adUseClient
                    
    dbexport.Execute "DROP DATABASE if exists " & dbserver
    dbexport.Execute "CREATE DATABASE " & dbserver
    
    'mysqldump -u root -proot main > "C:\MySQLBackups\main_22@AM.sql"
    
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & serveraddress & " " & dbserver & " <" & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)
    
    dbexport.Close
    Set dbexport = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Synced successfully", , "EzBiz"
    Exit Function
    
    '1. transmast
    TABLE_NAME = "transmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `transmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `transmast`"

    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & serveraddress & " " & dbserver & " <" & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '2. ASTRXMAST
    TABLE_NAME = "astrxmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `astrxmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `astrxmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '3. TRXEXPMAST
    TABLE_NAME = "trxexpmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxexpmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxexpmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & ""
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '4. trxexp_mast
    TABLE_NAME = "trxexp_mast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxexp_mast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxexp_mast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '5. trxincmast
    TABLE_NAME = "trxincmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxincmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxincmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '6. trxmast
    TABLE_NAME = "trxmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '7. returnmast
    TABLE_NAME = "returnmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `returnmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `returnmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '8. rtrxfile
    TABLE_NAME = "rtrxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `rtrxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `rtrxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '9. crdtpymt
    TABLE_NAME = "crdtpymt" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `crdtpymt` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `crdtpymt`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '10. dbtpymt
    TABLE_NAME = "dbtpymt" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `dbtpymt` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `dbtpymt`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '11. actmast
    TABLE_NAME = "actmast" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `actmast` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `actmast`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '12. cashatrxfile
    TABLE_NAME = "cashatrxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `cashatrxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `cashatrxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "


    '13. bank_trx
    TABLE_NAME = "bank_trx" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `bank_trx` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `bank_trx`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "

    '14. trxfile
    TABLE_NAME = "trxfile" & customercode
    db.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    db.Execute "CREATE TABLE " & TABLE_NAME & " LIKE `trxfile` "
    db.Execute "INSERT INTO " & TABLE_NAME & " SELECT * FROM `trxfile`"
    DoEvents
    cmd = Chr(34) & App.Path & "\mysqldump.exe" & Chr(34) & " -uroot -p###%%database%%###ret -h" & DBPath & " --routines --comments " & dbase1 & " " & TABLE_NAME & " > " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Execute "DROP TABLE if exists " & TABLE_NAME & " "
    DoEvents
    cmd = Chr(34) & App.Path & "\mysql.exe" & Chr(34) & " -u root -###%%database%%###ret -h " & serveraddress & " root_ezbiz < " & App.Path & "\Backup\dump.txt"
    Call execCommand(cmd)

    dbexport.Close
    Set dbexport = Nothing
    Screen.MousePointer = vbNormal
    MsgBox "Synced successfully", , "EzBiz"
    Exit Function
ERRHAND:
    Screen.MousePointer = vbNormal
    If err.Number = -2147467259 Then
        MsgBox "Connection failed. Make sure the Remote Server is live", , "EzBiz"
    Else
        MsgBox err.Description, , "EzBiz"
    End If
    
    
'    On Error Resume Next
'    db.RollbackTrans
End Function

'
'
'    '//Create Report Heading
'    '//chr(27) & chr(71) & chr(14) - for Enlarge letter and bold
'    '//chr(27) & chr(45) & chr(1) - for Enlarge letter and bold
'
'
'    'Print #1, Chr(27) & Chr(48) & Chr(27) & Chr(106) & Chr(216) & Chr(27) & _
'            Chr (106) & Chr(216) & Chr(27) & Chr(67) & Chr(60) & Chr(27) & Chr(80)
'    'Print #1, Chr(13)

Public Function DownloadFile(sURLFile As String, sLocalFilename As String) As Boolean
    Dim lRetVal As Long
    lRetVal = URLDownloadToFile(0, sURLFile, sLocalFilename, 0, 0)
    If lRetVal = 0 Then DownloadFile = True
End Function
