'#################################
'OBTENIR INFORMACIO EQUIP REMOT
'#################################'

'######################################################################
'IMPORTANT MODIFICAR DELS ALTRES SCRIPTS EN L'EXECUCIÓ DE LA CONSULTA:

	' wmiService ---> objSWbemServices '
	'Set wmiItemsDiskDrive = objSWbemServices.ExecQuery("SELECT * FROM ")
'######################################################################

'***************************************************************
'ESCRIU EN FITXER .\INFO.TXT TOTA LA INFORMACIO
'***************************************************************


Const ForAppending = 8
Dim strLogFile, strDate

strLogFile = ".\info.txt"

Set objFSO = CreateObject("Scripting.FileSystemObject")
if objFSO.FileExists(strLogFile) Then
	objFSO.DeleteFile(strLogFile)
End if
'Set objLogFile = objFSO.OpenTextFile(strLogFile, ForAppending, True)
'***************************************************************
'FI ESCRIU EN FITXER
'***************************************************************

Dim mess, strComputer, item
Dim objSWbemServices, objSWbemLocator

'********************************************
'CONNEXIO A EQUIP REMOT A: Root\CIMv2
'********************************************

StrComputer = InputBox("Indica l'equip GCB a analitzar (IP / Hostname) ")	
strUser = InputBox("Usuari (administrador de l'equip: ")											 
strPassword = InputBox ("Contrasenya: ")
Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")
Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, _
                                                     "Root\CIMv2", _
                                                     strUser, _
													 strPassword)
													
'********************************************

'*************************************************************
'INFORMACIO DE SISTEMA
'*************************************************************
Dim wmiItemsOSinfo
Set wmiItemsOSinfo = objSWbemServices.ExecQuery("SELECT * FROM Win32_OperatingSystem")		''Informació sistema operatiu
mess = mess & "-----------------------------------------------------------" & VBCrlf
mess = mess & VBCrlf
For each item in wmiItemsOSinfo
	With item
	mess = mess & " HOSTNAME:	" & .CSName & VBCrlf
	mess = mess & VBCrlf
	mess= mess & " Sistema Operatiu:	" & .Caption & " " & .OSArchitecture & VBCrlf
	mess = mess & VBCrlf
	End With
Next

'LLEGIR REGISTRE  (PLATAFORMA (DEPARTAMENT/ORGANISME)
'***************************************************************

Const HKEY_LOCAL_MACHINE = &H80000002

Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _
    strComputer & "\root\default:StdRegProv")

'CONSULTA HKEY_LOCAL_MACHINE\SOFTWARE\PLATAFORMA\INFORMACION_PUESTO ---> Departament -- Organisme
strKeyPath = "SOFTWARE\Plataforma\Informacion_Puesto"
strDepartament = "Departament"
strOrganisme = "Organisme"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strDepartament,strValue
mess = mess & "     >> Departament: " & strValue & VBCrlf
mess= mess & VBCrlf
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strOrganisme,strValue
mess = mess & "     >> Organisme: " & strValue & VBCrlf
mess = mess & VBCrlf

'CONSULTA HKEY_LOCAL_MACHINE\SOFTWARE\PLATAFORMA\MASTER ---> Display Version
strKeyPath = "SOFTWARE\Plataforma\MASTER"
strVersioMaqueta = "DisplayVersion"

oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strVersioMaqueta,strValue
'strValue = right(strValue,5)
mess = mess & "     >> Versio maqueta: " & strValue & VBCrlf
mess= mess & VBCrlf

'***************************************************************

''INFORMACIO USUARI LOGUEJAT
''***************************************************************
ComputerName ="."
Set wmiServiceusuari = GetObject("winmgmts:\\" & StrComputer)
Dim wmiItemusuari
Set wmiItemusuari = wmiServiceusuari.ExecQuery("SELECT * FROM Win32_ComputerSystem")
For Each item in wmiItemusuari
	With item
		mess = mess & " USUARI LOGUEJAT:	" & .UserName & VBCrlf
		mess = mess & VBCrlf
	End With
Next




Dim wmiItemsSysinfo
Set wmiItemsSysinfo = objSWbemServices.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct")		'Informacio de l'equip
For Each item in wmiItemsSysinfo
	With item
		mess = mess & " MODEL:	" & .Vendor & " " & .Name & VBCrlf
		mess = mess & VBCrlf
		mess = mess & " S/N:	" & .IdentifyingNumber & VBCrlf
	End With
Next

mess = mess & VBCrlf

'INFORMACIO CPU
'***************************************************************
ComputerName = "."
Set wmiService = GetObject("winmgmts:\\" & StrComputer)
Set wmiItems = wmiService.ExecQuery("SELECT * FROM win32_Processor")
For Each item in wmiItems
	With item
		mess = mess & " CPU:	" & .Name & .AddressWidth & VBCrlf
		mess = mess & VBCrlf
	End With
Next


'INFORMACIO RAM
'***************************************************************
Dim wmiItemsMemory
Set wmiItemsMemory = objSWbemServices.ExecQuery("SELECT * FROM win32_PhysicalMemory")
For Each item in wmiItemsMemory
	With item
		'mess = mess & "- Modul instalat a: " & .BankLabel & " "  & .DeviceLocator &  VBCrlf & "- Capacitat: " & left(.Capacity/1024^3,6) & " GB " & VBCrlf & "- Fabricant: " & .ManuFacturer & VBCrlf & "- Num.serie: " & .PartNumber & VBCrlf
		mess = mess & " RAM:	" & left(.Capacity/1024^3,6) & " GB " & VBCrlf
		mess = mess & VBCrlf
	End With
Next
'INFORMACIO DISC
'***************************************************************
Dim wmiItemsDisk
Set wmiItemsDisk = objSWbemServices.ExecQuery("SELECT * FROM win32_LogicalDisk")
For Each item in wmiItemsDisk
	With item
		mess = mess & " DISC: Espai lliure:	" & .Caption & " " & left(.FreeSpace/1024^3,3) & " GB " & "/ " & left(.Size/1024,3) & " GB" & VBCrlf
		mess = mess & VBCrlf
	End With
Next

'***************************************************************
'INFORMACIO XARXA
'***************************************************************
mess = mess & VBCrlf
Set wmiItemsxarxa = objSWbemServices.ExecQuery("SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
mess = mess & "INFORMACIO DE XARXA " & VBCrlf 
mess = mess & "------------------------" & VBCrlf
mess = mess & VBCrLf
For Each item in wmiItemsxarxa
	With item
		mess = mess & "ADAPTADOR DE XARXA: " & .Description & VBCrLfd
		mess = mess & VBCrlf  
		For Each strIPSubnet in .IPSubnet
			subxarxa= subxarxa & strIPSubnet
			'*******************************************************************************************************************************
			'Retallem els 2 últims digits que agafa de la Mascara de subxarxa (64) per deixar unicament els 12 digits vàlids 255.255.255.255
			subxarxa = Left(subxarxa,15)
		Next
		For Each strIPAddress in .IPaddress 
			mess = mess & VBCrlf  
			mess = mess & "     >> IP: " & strIPAddress & " /  " & subxarxa & VBCrLf 
			mess = mess & VBCrlf
			mess = mess & "     >> MAC: " & .MACAddress & VBCrLf
			mess = mess & VBCrlf
		Next
		mess = mess & "     >> DHCP Habilitat: " & .DHCPEnabled & VBCrlf
			mess = mess & VBCrLf
	End With
Next
mess = mess & "-----------------------------------------------------------" & VBCrlf
mess = mess & VBCrLf

'----------------------------------------------------------------------




'INFORMACIO PANTALLA PER POWERSHELL '----( executa el powershell info_monitor.ps1)  --------------------------------------------------------------------
mess = mess & "PANTALLA" & VBCrlf
mess = mess & "------------------------" & VBCrlf

'HABILITA FUNCIO D'EXECUCIO D'SCRIPTS POWERSHELL EN EQUIP QUE EXECUTA SCRIPT (local)
strInstruccio0 = "powershell.exe Set-ExecutionPolicy Unrestricted"
strDosCommand0 = strInstruccio0
Set objShell = CreateObject("Wscript.Shell")
Set objExec = objShell.Exec(strDOSCommand0)
strPSResults0 = objExec.StdOut.ReadAll




strInstruccio = "powershell -file info_monitor.ps1"
strDOSCommand = strInstruccio & " " & strComputer 	'Passa com a parametre d'execució de l'script powershell(.ps1) el host a analitzar

'strDOSCommand = "powershell -file -paramter info_monitor.ps1"
Set objShell = CreateObject("Wscript.Shell")
Set objExec = objShell.Exec(strDOSCommand)
strPSResults = objExec.StdOut.ReadAll
'WScript.Echo(strPSResults)
mess = mess & strPSresults
'--------------------------------------------------------------------------------------------------------------------------------------------------------

' FI - ESCRIU PER PANTALLA TOTA LA INFORMACIO OBTINGUDA
'Esriu tota la info del recurregut en fitxer
WScript.Echo mess
strEscriu = mess


