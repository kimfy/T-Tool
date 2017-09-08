;T-Tool
#RequireAdmin
; Kommenter ut #RequireAdmin for å ikke bli spurt om admin under testing. NB må være med på release
#include <Icons.au3>
#include <Date.au3>
#include <File.au3>
#include <Misc.au3>
#include <GuiTab.au3>
#include <WinAPI.au3>
#include <WinAPIFiles.au3>
#include <Constants.au3>
#include <ComboConstants.au3>
#include <EditConstants.au3>
#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <InetConstants.au3>
#include <MsgBoxConstants.au3>
#include <WinAPIFiles.au3>
#include <Array.au3>
#include <EventLog.au3>


AutoItSetOption("WinTitleMatchMode", 2)     ; this allows partial window titles to be valid!

;To-do list'
; Fikse Ping av DC til å inkludere flere DCer
; Fikse delete temp da den pt. Kun er copy paste fra forum, Startet på ny kode i 1.5.4.4
; Finne bruker på PC om bruker ikke har lokal admin, mulig fikset i 1.5.4.3
; ATS melder om at hostfil inneholder ip selv om den ikke gjør det. Trolig fikset i 1.5.3.6


;Sandbox
Dim $Sandbox = "No"
Dim $TestKnapp = 0 ; Sett verdi til 1 for å vise knappen

;Decl.
Dim $SWName = "T-Tool"
Dim $SWVersion = "1.5.4.8"
Dim $SWStable = "Unstable"
Dim $i
Dim $Sandbox
Dim $PingIP
Dim $PingStatus
Dim $RandomTXT[100]
Dim $RandomTotal
Dim $RandomTest
Dim $aNumber
Dim $WUAFILE
Dim $DTVFilePath
Dim $DTVFile
Dim $Link
Dim $ScreenX = @DesktopWidth
Dim $ScreenY = @DesktopHeight
Dim $MailAddress = "olai.wanvik@torghatten.no"
Dim $ShortcutFile
Dim $Portable
Dim $Hostname = @ComputerName
Dim $NSIP
Dim $BetaCheck
Dim $Safemode
Dim $GPTag
Dim $FileExtStatus = ""
Dim $FileExtStatusTag = ""
Dim $HiddenFolderStatus = ""
Dim $HiddenFolderStatusTag = ""
Dim $DNSReadLine
Dim $DNSPrim
Dim $DNSSec
Dim $DirUser
Dim $RunInput

;GUI Variabler
Dim $GUIStatus = 0 ; $GUIStatus settes til 1 i Post-GUI
Dim $GUIWidth = 465 ;Default 430
Dim $GUIHeight = 305 ;Default 305
Dim $SNWidth = 185 ;Default 185, men endres i egen kode ut ifra lengde på SN
Dim $LabelLeft = 245 ;Default 210
Dim $LabelTop = 5 ;Default 5
Dim $LabelSpace = 20 ; - Setter avstanden mellom labels
Dim $LabelWidth = 250 ;Default 250
Dim $LabelHeight = 20 ;Default 20
Dim $BTN_Left_Row1 = 5 ;Setter avstand fra venstre side på knapper i rad 1
Dim $BTN_Top = 35
Dim $BTN_Width = 113
Dim $BTN_Height
Dim $BTN_Left_Row2 = $BTN_Width + 10 ;
Dim $BTW_Widht_Small

Global $g_idMemo

;Decl. AddTrustedSite
Dim $Username
Dim $Password
Dim $intracred = "CMDKEY /add:intra.torghatten.no /user:torghatten\"&$Username&" /pass:"&$Password
Dim $hosts
Dim $System32 = StringLeft(@SystemDir,11)&"System32\"
Dim $line = 0
Dim $string
Dim $EditHosts

;Del temp
Dim $Debug
Dim $AllFiles
Dim $FolderToDelete
Dim $crt
Dim $delete

;HostInventory
Dim $SN
Dim $Minne
Dim $Systemtype = @OSArch
Dim $WinVersjon = @OSVersion
Dim $WinRelease
Dim $WinReleaseID
Dim $Vendor
Dim $VendorVersion
Dim $ComputerModelExtended = 0

Dim $PublicIP
Dim $DCPing
Dim $LastBootTime
Dim $DeviceInDomain
Dim $DeviceDomainStatus

;Pre GUI Funcs
ProgressOn($SWName,"","Running pre-GUI functions")
ProgressSet(10,"Getting system info")
GetSysFunc()
ProgressSet(12,"Getting serialnumer")
GetSNFunc()
ProgressSet(14,"Getting RAM")
GetRAMFunc()
ProgressSet(16,"Getting ReleaseID")
GetReleaseIDFunc()
ProgressSet(18,"Checking portablemode")
PortableCheck()
ProgressSet(20,"Adjusting GUI, pre build")
WideSN()
ProgressSet(22,"Checking for domain")
CheckDeviceDomainStatus()
ProgressSet(24,"Checking computermodel")
GetComputerModel()
ProgressSet(26,"Getting hidden folders and files status")
ShowHideExt()
ShowHideFolders()

;RandomTXT
$RandomTXT[0] = "Henter og setter sammen IKEA-møbler"
$RandomTXT[1] = "Høster og sår parametre og variabler"
$RandomTXT[2] = "Samler resultater fra prøver som ikke er tatt"
$RandomTXT[3] = "Undersøker hva som ligger under denne boksen"
$RandomTXT[4] = "Forsker på spørsmål og svar om livet"
$RandomTXT[5] = "Kalkulerer kvantefysikk i tredje grad"
$RandomTXT[6] = "Ser etter nye GIFs på 9GAG"
$RandomTXT[7] = "Det kan du sikkert lure på"
$RandomTXT[8] = "01101000 01100001 01101000 01100001 00001101 00001010"
$RandomTXT[9] = Chr(34)&"Bad joke"&Chr(34)
$RandomTXT[10] = "Sletter C:\Windows\System32"
$RandomTXT[11] = "Vasker og pusser skrivebordsbakgrunnen"
$RandomTXT[12] = "Kjøper hest på bruktbilmarked"
$RandomTXT[13] = "Plukker epler fra apelsintrær"
$RandomTXT[14] = "Kan ikke se å finne denne teksten"
$RandomTXT[15] = "Bar barbar-bar-barbar bar bar barbar-bar-barbar"
$RandomTXT[16] = "Prøver å fortsette der noen slapp"
$RandomTXT[17] = "Sjekker at strømmen står i"
$RandomTXT[18] = "Slår av og på"
$RandomTXT[19] = "Loading Sky Factory 3"
$RandomTotal = 19

ProgressSet(40,"Building GUI")
#Region ### START Koda GUI section ### Form=
$TTool = GUICreate($SWName&" - v."&$SWVersion&$Portable, $GUIWidth, $GUIHeight)
GUISetBkColor(0xF7F7F7)

;Tools -
$MenuTools = GUICtrlCreateMenu("Tools")
$MenuDisableFastStartup = GUICtrlCreateMenuItem("Disable fast startup", $MenuTools)
$MenuDeleteCreds = GUICtrlCreateMenuItem("Delete Microsoft credentials (experimental!)",$MenuTools)
$MenuFlushDNSCache = GUICtrlCreateMenuItem("Flush cache", $MenuTools)
$MenuSystem = GUICtrlCreateMenu("System",$MenuTools)
$MenuRenewIP = GUICtrlCreateMenuItem("Renew IP",$MenuTools)
$MenuGPResult = GUICtrlCreateMenuItem("GPResult",$MenuTools)
$MenuOutlookSafemode = GUICtrlCreateMenuItem("Start Outlook i Safemode",$MenuTools)
$MenuResyncTime = GUICtrlCreateMenuItem("Resync klokke",$MenuTools)
$MenuGodMode = GUICtrlCreateMenuItem("GodMode",$MenuTools)
$MenuGetDNS = GUICtrlCreateMenuItem("Get DNS",$MenuTools)
$MenuRepOffice = GUICTrlCreateMenuItem("Office 365 Reparasjon",$MenuTools)
$MenuCMD = GUICtrlCreateMenuItem("CMD",$MenuTools)

;Tools - System -
$MenuShutdown = GUICtrlCreateMenuItem("Shutdown",$MenuSystem)
$MenuReboot = GUICtrlCreateMenuItem("Reboot",$MenuSystem)
$MenuLastBootTime = GUICtrlCreateMenuItem("Get last boot time",$MenuSystem)
$MenuInstallDate = GUICtrlCreateMenuItem("Get OS install date",$MenuSystem)
$MenuGetComputerModel = GUICtrlCreateMenuItem("Get computermodel",$MenuSystem)
$MenuResolution = GUICtrlCreateMenuItem("Resolution: "&$ScreenX&"x"&$ScreenY,$MenuSystem)
GUICtrlSetState(-1,$GUI_DISABLE)

;Controls -
$MenuControls = GUICtrlCreateMenu("Controls")
$MenuPowerCfg = GUICtrlCreateMenuItem("Strømalternativer",$MenuControls)
$MenuAppWiz = GUICtrlCreateMenuItem("Legg til eller fjern programmer", $MenuControls)
$MenuInet = GUICtrlCreateMenuItem("Alternativer for internett",$MenuControls)
$MenuLusrmgr = GUICtrlCreateMenuItem("Lokale brukere og grupper",$MenuControls)
$MenuDeviceManager = GUICTrlCreateMenuITem("Enhetsbehandling", $MenuControls)
$MenuEventViewer = GUICtrlCreateMenuItem("Event viewer",$MenuControls)
$MenuTaskScheduler = GUICtrlCreateMenuItem("Planlagte oppgaver",$MenuControls)
$MenuCredentialManager = GUICtrlCreateMenuItem("Legitimasjonsbehandling",$MenuControls)
$MenuResMon = GUICtrlCreateMenuItem("Ressursovervåkning",$MenuControls)
$MenuTaskMgr = GUICtrlCreateMenuItem("Oppgavebehandling",$MenuControls)
$MenuDevicesAndPrinters = GUICtrlCreateMenuItem("Enheter og skrivere",$MenuControls)
$MenuRegistry = GUICtrlCreateMenuItem("Registry",$MenuControls)
$MenuWinUpdate = GUICtrlCreateMenuItem("Windows Update",$MenuControls)
$MenuWinFirewall = GUICtrlCreateMenuItem("Windows brannmur",$MenuControls)

;Links
$MenuLinks = GUICtrlCreateMenu("Links")
$MenuLinkSS = GUICtrlCreateMenuitem("Selfservice - Pureservice 2.5",$MenuLinks)
$MenuLinkEyeshare = GUICtrlCreateMenuItem("Eyeshare",$MenuLinks)
$MenuLinkHPWarranty = GUICtrlCreateMenuItem("Check HP Warranty",$MenuLinks)
$MenuLinkEK = GUICtrlCreateMenuItem("Portalen / EKWeb",$MenuLinks)
$MenuLinkOffice365 = GUICtrlCreateMenuItem("Office365 login",$MenuLinks)
$MenuLinkTeamviewerqs = GUICtrlCreateMenuItem("Teamviewer QS",$MenuLinks)
$MenuLinkTeamviewerhost = GUICtrlCreateMenuItem("Teamviewer Host",$MenuLinks)

;Links - Create
$MenuLinksCreate = GUICtrlCreateMenu("Create bookmark on desktop",$MenuLinks)
$MenuLinksCreateEyeshare = GUICtrlCreateMenuitem("Eyeshare",$MenuLinksCreate)
$MenuLinksCreateEK = GUICtrlCreateMenuitem("Portalen / EKWeb",$MenuLinksCreate)

;Help -
$MenuHelp = GUICtrlCreateMenu("&Help")
$MenuChangelog = GUICtrlCreateMenuItem("Open changelog", $MenuHelp)
$MenuReportError = GUICtrlCreateMenuItem("Rapporter feil", $MenuHelp)
$MenuKillTTool = GUICTrlCreateMenuItem("Kill T-Tool", $MenuHelp)
$MenuSafemode = GUICtrlCreateMenuItem("Safemode",$MenuHelp)
$MenuCurrentV = GuiCtrlCreateMenuItem("Current version: "&$SWVersion,$MenuHelp)
GUICtrlSetState(-1,$GUI_DISABLE)

;Buttons
$BTN_GPUPDATE = GUICtrlCreateButton("GPUpdate", $BTN_Left_Row1, 5, $BTN_Width, 35)
$BTN_ATS = GUICtrlCreateButton("AddTrustedSite", $BTN_Left_Row1, 45, $BTN_Width, 35)
$BTN_PING = GUICtrlCreateButton("Ping", $BTN_Left_Row1, 85, $BTN_Width, 35)
$BTN_WUAU = GUICtrlCreateButton("Windows Update Fix", $BTN_Left_Row1, 125, $BTN_Width, 35)
$BTN_RUN = GUICtrlCreateButton("Run", $BTN_Left_Row1, 165, $BTN_Width, 35)
$BTN_ResMon = GUICtrlCreateButton("Ressursovervåkning", $BTN_Left_Row1, 205, $BTN_Width, 35)
$BTN_TaskMgr = GUICtrlCreateButton("Oppgavebehandling", $BTN_Left_Row1, 245, $BTN_Width, 35)
$BTN_OpprettSak = GUICtrlCreateButton("Opprett sak",$BTN_Left_Row2,5,$BTN_Width,35)
$BTN_DelTemp = GUICtrlCreateButton("Slett temp",$BTN_Left_Row2,45,$BTN_Width,35)
$BTN_OpenHosts = GUICtrlCreateButton("Åpne Hosts",$BTN_Left_Row2,85,$BTN_Width,35)
$BTN_KillOneDrive = GUICtrlCreateButton("Kill OneDrive",$BTN_Left_Row2,125,$BTN_Width,35)
$BTN_FileExt = GUICtrlCreateButton($FileExtStatusTag&" filtyper",$BTN_Left_Row2,165,$BTN_Width,35)
$BTN_HiddenFolder = GUICtrlCreateButton($HiddenFolderStatusTag&" skjulte mapper",$BTN_Left_Row2,205,$BTN_Width,35)
$BTN_Test = GUICtrlCreateButton("Test",$BTN_Left_Row2,245,$BTN_Width,35)


;HostInvGUI
$InventoryLabel = GUICtrlCreateLabel("Enhetsinformasjon: ", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
GUICtrlSetFont(-1, 9, 800)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace ; Denne regner ut hvor neste label skal stå -----------

$PcNavn = GUICtrlCreateLabel("PC Navn: "&@ComputerName, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$ComputerType = GUICtrlCreateLabel("PC type: "&$Vendor&" "&$VendorVersion, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$IPAddress = GUICtrlCreateLabel("IP Adresse: ", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$SystemType = GUICtrlCreateLabel("Systemtype: "& $WinVersjon & " " & $Systemtype, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$WinRelease = GUICtrlCreateLabel("Versjon: "&$WinReleaseID, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$MinneLabel = GUICtrlCreateLabel("Installert Minne: "&$Minne&" GB", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$Serienummer = GUICtrlCreateLabel("Serienummer: "&$SN, $LabelLeft, $LabelTop, $SNWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$UsernameLabel = GUICtrlCreateLabel("Brukernavn: "&@UserName, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$DeviceInDomain = GUICtrlCreateLabel($DeviceDomainStatus, $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$DCStatus = GUICtrlCreateLabel("", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$LastBootTimeLabel = GUICtrlCreateLabel("", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
$LabelTop = $LabelTop + $LabelSpace

$OSInstallDateLabel = GUICtrlCreateLabel("", $LabelLeft, $LabelTop, $LabelWidth, $LabelHeight)
;GUICtrlSetFont(-1, 9)
GUICtrlSetColor(-1, 0x181818)
GUICtrlSetBkColor(-1, 0xF7F7F7)
$LabelTop = $LabelTop + $LabelSpace

;Det er ikke plass til flere linjer med label uten å endre på GUI.
;OSInstallDate er 1px inn i IT-TORG logoen pt.

GUICtrlCreatePic(@ScriptDir&"\sys\GUI\ITTORG.jpg",$GUIWidth - 100,$GUIHeight - 67,100,63)
GUISetIcon(@ScriptDir&"\sys\gui\icon.ico")
#NoTrayIcon
GUISetState(@SW_SHOW)

#Region Hotkeys. "+" = shift, "!" = alt, "^" = ctrl,
;Dim $AccelTable[1][2] = [["!p", $BTN_PING]]
Dim $AccelTable[1][2] = [["!q", $BTN_TEST]]
GUISetAccelerators($AccelTable)
#EndRegion

;Testknapp
GUICtrlSetState($BTN_Test,$GUI_HIDE)
GUICtrlSetBKColor($BTN_Test, $COLOR_RED)
GUICtrlSetColor($BTN_Test, $COLOR_WHITE)
GUICtrlSetFont($BTN_Test, 15)
If $TestKnapp = 1 Then GUICtrlSetState($BTN_Test, $GUI_SHOW)
#EndRegion ### END Koda GUI section ### --------------------------------------------------------------------------

;Post GUI Funcs
ProgressOn($SWName,"","Running Post-GUI functions")
ProgressSet(60,"Running post GUI functions")
$GUIStatus = 1
ProgressSet(65,"Checking for domaincontrollers")
PingTHDCFunc()
ProgressSet(70,"Getting IP-address")
GetNSIP()
ProgressSet(75,"Setting safemode"&@CRLF&"(If available)")
SetSafemode()
SafemodeApps()
;Her lukkes loadingbar
ProgressSet(100,"All Done")
ProgressOff()

FindUser()



While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		 Case $GUI_EVENT_CLOSE
			Exit
		 Case $MenuDisableFastStartup
			DisableFastStartupFunc()
		 Case $MenuFlushDNSCache
			FlushDNSCacheFunc()
		 Case $MenuDeleteCreds
			Betamode()
			If $BetaCheck = "Yes" then DeleteCredsFunc()
		 Case $MenuLastBootTime
			LastBootTimeFunc()
		 Case $BTN_PING
			PingFunc()
		 Case $BTN_ATS
			AddTrustedSite()
		 Case $BTN_GPUPDATE
			GPUpdateFunc()
		 Case $MenuReboot
			RebootFunc()
		 Case $MenuShutdown
			ShutdownFunc()
		 Case $MenuChangeLog
			OpenChangelogFunc()
		 Case $MenuInstallDate
			InstallDateFunc()
		 Case $MenuRenewIP
			RenewIPFunc()
		 Case $MenuGPResult
			GPResultFunc()
		 Case $MenuKillTTool
			KillTTool()
		 Case $BTN_WUAU
			WindowsUpdateFixFunc()
		 Case $MenuRepOffice
			BetaMode()
			If $BetaCheck = "Yes" then RepairOffice365()
		 Case $MenuPowerCfg
			StromalternativerFunc()
		 Case $MenuAppWiz
			AppWizFunc()
		 Case $MenuInet
			INetFunc()
		 Case $MenuLusrmgr
			LusrmgrFunc()
		 Case $MenuDeviceManager
			DeviceManager()
		 Case $MenuEventViewer
			EventViewer()
		 Case $MenuTaskScheduler
			TaskScheduler()
		 Case $MenuLinkSS
			$Link="https://helpdesk.torghatten.no/selfservice"
			OpenLink()
		 Case $MenuLinkEyeshare
			$Link = "https://eyeshare.torghatten.no/eyeshare1"
			OpenLink()
		 Case $MenuLinkHPWarranty
			$Link = "https://support.hp.com/us-en/checkwarranty"
			OpenLink()
		 Case $MenuLinkEK
			$Link = "https://ekweb.fosen.no/ekweb/start.aspx?main=1"
			OpenLink()
		 Case $MenuLinkOffice365
			$Link = "https://portal.office.com"
			OpenLink()
		 Case $MenuLinkTeamviewerQS
			$Link = "https://get.teamviewer.com/torghattenqs"
			OpenLink()
		 Case $MenuLinkTeamviewerHost
			$Link = "https://get.teamviewer.com/torghattenhost"
			OpenLink()
		 Case $MenuOutlookSafemode
			StartOutlookInSafemode()
		 Case $MenuReportError
			ReportError()
		 Case $MenuLinksCreateEyeshare
			$ShortcutFile = "eye-share.website"
			BookmarkCreate()
		 Case $MenuLinksCreateEK
			$ShortcutFile = "Portalen.website"
			BookmarkCreate()
		 Case $MenuCredentialManager
			CredentialManager()
		 Case $BTN_TaskMgr
			TaskManager()
		 Case $MenuTaskMgr
			TaskManager()
		 Case $MenuResMon
			ResMon()
		 Case $BTN_ResMon
			ResMon()
		 Case $MenuDevicesAndPrinters
			DevicesAndPrinters()
		 Case $MenuRegistry
			RegEdit()
		 Case $MenuWinUpdate
			WinUpdate()
		 Case $MenuWinFirewall
			WinFirewall()
		 Case $MenuSafemode
			ChangeSafemode()
			SafemodeApps()
		 Case $MenuResyncTime
			ResyncTime()
		 Case $MenuGodMode
			GodMode()
		 Case $BTN_OpprettSak
			OpprettSak()
		 Case $BTN_DelTemp
			DelTemp()
		 Case $BTN_OpenHosts
			OpenHosts()
		 Case $BTN_KillOneDrive
			KillOneDriveEXE()
		 Case $MenuGetComputerModel
			$ComputerModelExtended = 1
			GetComputerModel()
		 Case $BTN_FileExt
			   ShowFileExt()
			   ShowHideExt()
		 Case $BTN_HiddenFolder
			   ShowHiddenFolder()
			   ShowHideFolders()
		 Case $MenuGetDNS
			   GetDNS()
		 Case $BTN_RUN
			   RunCommand()
		 Case $MenuCMD
			   CMD()

		 Case $BTN_Test
			;Bruk denne knappen til å teste funksjoner. Sett $TestKnapp til 1 for å vise knappen
			BetaMode()
			If $BetaCheck = "Yes" then ShowHiddenFolder()

	EndSwitch
WEnd

;Functions - Paste gui above here


Func GPUpdateFunc()
 Run("gpupdate /force")
 ;Sleep(300)
 WinWait("gpupdate.exe")
 Do
 $i = WinExists("gpupdate.exe")
 Until $i = 0
 MsgBox(0,$SWName,"GPUpdate fullført",2)

EndFunc

;AddTrustedSiteFunctions --------------------------------------------------------------------------------------------------------------------------------------------

Func GetCred()
$Username = InputBox(" ","Skriv inn brukernavn"&@CRLF&@CRLF&"NB! IKKE skriv inn domene (torghatten)"&@CRLF&@CRLF&"Brukernavn lokalt på denne maskina er: "&@UserName&@CRLF&"Det trenger ikke å være samme brukernavn i AD.")
$Password = InputBox(" ","Skriv inn passord")
If @error = 1 Then Exit
EndFunc

Func CheckCred()
$i = MsgBox(36,"","Er brukernavn og passord korrekt?"&@crlf&@crlf&"Brukernavn: "&$Username&@crlf&"Passord: "&$Password)
EndFunc

Func RegCreds()

   ;MsgBox(64,"","Adding credidentials")
   $intracred = "CMDKEY /add:intra.torghatten.no /user:torghatten\"&$Username&" /pass:"&$Password
  ;MsgBox(0,"",$intracred)
   RunWait(@ComSpec & " /c " & $intracred)
   $intracred = "CMDKEY /add:eyeshare.torghatten.no /user:torghatten\"&$Username&" /pass:"&$Password
   RunWait(@ComSpec & " /c " & $intracred)
   $intracred = "CMDKEY /add:ekweb.fosen.no /user:torghatten\"&$Username&" /pass:"&$Password
   RunWait(@ComSpec & " /c " & $intracred)
EndFunc

Func AddIntranet()

   ;MsgBox(64,"","Writing to registry")
   RegWrite("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\torghatten.no\intra","*", "REG_DWORD","1")
   RegWrite("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\torghatten.no\eyeshare","*", "REG_DWORD","1")
   RegWrite("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\intra.thn.no","*", "REG_DWORD","1")
   RegWrite("HKEY_CURRENT_USER\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\ZoneMap\Domains\ekweb.fosen.no","*", "REG_DWORD","1")
EndFunc

Func OpenURL()
ShellExecute("https://intra.torghatten.no/ekweb/delta.aspx?smp=1")
EndFunc

Func EditHosts()
 $EditHosts = MsgBox(36,"","Skal intra.torghatten.no legges inn i hosts-fila?"&@CRLF&"Det skal kun gjøres på fartøy som ikke har DigiRuter")
If $EditHosts = 6 Then
;MsgBox(0,"","Yes")
	  ;$hosts = Run("notepad.exe "& $System32& "Drivers\etc\hosts")
   $line = 0
	  Do
		 $line = $line + 1
		 $string =  FileReadLine($System32&"Drivers\etc\hosts",$line)
			If StringInStr($string,"intra.torghatten.no") Then
			   MsgBox(0,"","Oppføringen for intra.torghatten.no finnes allerede i hosts på line "&$line)
			   $line = 200
			Else
				  If $line = 199 Then
					 ;MsgBox(0,"","String not found"&@CRLF&"Line is "&$line)
					 RunWait(@ComSpec & " /c " & "echo 185.125.160.250	intra.torghatten.no >> "&$System32&"Drivers\etc\hosts")
				  Else
				  EndIf
			EndIf
	  Until $line = 200
Else
   ;MsgBox(0,"","No")
EndIf

EndFunc

Func EditHostsF()
    $line = 0
	  Do
		 $line = $line + 1
		 $string =  FileReadLine($System32&"Drivers\etc\hosts",$line)
			If StringInStr($string,"intra.torghatten.no") Then
			   MsgBox(0,"","Oppføringen for intra.torghatten.no finnes allerede i hosts på line "&$line)
			   $line = 200
			Else
				  If $line = 199 Then
					 ;MsgBox(0,"","String not found"&@CRLF&"Line is "&$line)
					 RunWait(@ComSpec & " /c " & "echo 185.125.160.250	intra.torghatten.no >> "&$System32&"Drivers\etc\hosts")
				  Else
				  EndIf
			EndIf
	  Until $line = 200
   EndFunc

;End of AddTrustedSiteFunc------------------------
Func AddTrustedSite()

Do
   GetCred()
   CheckCred()
If $i = 7 Then
   GetCred()
Else
   RegCreds()
   AddIntranet()
   EditHostsF()
EndIf

Until $i = 6

;RegCreds()
;AddIntranet()
;EditHostsF()

OpenURL()
EmptyIVAR()
EndFunc


;HostInventoryFuncs
Func GetSysFunc()
   If $Systemtype = "X64" Then
	  $Systemtype = "64bit"
   Else
	  $Systemtype = "32bit"
   EndIf

   If $WinVersjon = "WIN_7" Then $WinVersjon = "Windows 7"
	  If $WinVersjon = "WIN_XP" Then $WinVersjon = "Windows XP"
		 If $WinVersjon = "WIN_VISTA" Then $WinVersjon = "Windows Vista"
			   If $WinVersjon = "WIN_10" Then $WinVersjon = "Windows 10"
				  If $WinVersjon = "WIN_8.1" Then $WinVersjon = "Windows 8.1"
					 If $WinVersjon = "WIN_8" Then $WinVersjon = "Windows 8"

EndFunc

Func GetSNFunc()

If FileExists(@tempdir&"\serienummer.txt") Then
   $SN = FileReadLine(@tempdir&"\serienummer.txt",2)
   $SN = StringStripWS($SN, $STR_STRIPALL)
Else
   runWait(@comSpec & ' /c wmic bios get serialnumber > "' & @tempDir & '\serienummer.txt"',@ScriptDir,@SW_HIDE)

EndIf

If FileExists(@tempdir&"\serienummer.txt") Then
   $SN = FileReadLine(@tempdir&"\serienummer.txt",2)
   $SN = StringStripWS($SN, $STR_STRIPALL)
Else
   GetSNFunc()
EndIf
EndFunc

Func GetRAMFunc()
  $Minne = MemGetStats()
	  $Minne = $Minne[1]/1024000
	  $Minne = StringLeft($Minne,4)
$Minne = Round($Minne)
   EndFunc

Func GetReleaseIDFunc()
$WinReleaseID = RegRead("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion", "ReleaseId")
If $WinVersjon = "Windows 7" Then $WinReleaseID = "Finnes ikke på "&$WinVersjon
   If $WinVersjon = "Windows Vista" Then $WinReleaseID = "Finnes ikke på "&$WinVersjon
	  If $WinVersjon = "Windows XP" Then $WinReleaseID = "Finnes ikke på "&$WinVersjon
		 ;If $WinVersjon = "Windows 10" Then $WinReleaseID = "Finnes ikke på "&$WinVersjon
EndFunc

Func PingTHDCFunc()
   ;Edit 17.07.2017, set colors
	  ;Edit 25.08.2017, lagt til ping av brannmur if all else fail
$DCPing = Ping("10.40.101.101")

If $DCPing = 0 Then $DCPing = Ping("10.40.101.1") + 2

If $DCPing = 1 Then
   GUICtrlSetData($DCStatus,"Har kontakt med DC i Torghatten")
   GUICtrlSetColor($DCStatus, $COLOR_GREEN)
Else
   GUICtrlSetData($DCStatus,"Har ikke kontakt med DC")
   GUICtrlSetColor($DCStatus, $COLOR_RED)
EndIf

If $DCPing = 3 Then
   GUICtrlSetData($DCStatus,"Maskina er i Torghatten nettverket")
   GUICtrlSetColor($DCStatus, $COLOR_GREEN)
EndIf
  ; MsgBox(0,"",$DCPing)

EndFunc

Func LastBootTimeFunc()
   ProgressOn($SWName,"")
   ProgressSet(20)
   ProgressSet(50)
   $i = Random(0,$RandomTotal)
   ProgressSet(80,$RandomTXT[$i])
   runWait(@comSpec & ' /c systeminfo > ' & @TempDir & '\boottime.txt"',@ScriptDir,@SW_HIDE)
   $LastBootTime = FileReadLine(@TempDir & "\boottime.txt",12)
   $LastBootTime = StringRight($LastBootTime,20)
      $i = Random(0,$RandomTotal)
   ProgressSet(90,$RandomTXT[$i])
   GUICtrlSetData($LastBootTimeLabel,"Last reboot: "&$LastBootTime)
   ProgressSet(100)
   ProgressOff()
   MsgBox(0,$SWName,"Last reboot: "&$LastBootTime)
EndFunc

Func InstallDateFunc()
   ProgressOn($SWName,"")
   ProgressSet(20)
   ProgressSet(50)
   $i = Random(0,$RandomTotal)
   ProgressSet(80,$RandomTXT[$i])
   runWait(@comSpec & ' /c systeminfo > ' & @tempDir & '\systeminfo.txt"',@ScriptDir,@SW_HIDE)
   $InstallDate = FileReadLine(@tempDir & "\systeminfo.txt",11)
   $Installdate = StringRight($InstallDate,20)
   $Installdate = StringLeft($InstallDate,10)
      $i = Random(0,$RandomTotal)
   ProgressSet(90,$RandomTXT[$i])
   GUICtrlSetData($OSInstallDateLabel,"Install date: "&$InstallDate)

   ProgressSet(100)
   ProgressOff()
   MsgBox(0,$SWName,"OS Install date: "&$InstallDate)
EndFunc

Func DisableFastStartupFunc()
   ProgressOn($SWName,"Deaktiverer rask oppstart")
   ProgressSet(50,"Deaktiverer rask oppstart")
$CMD = "powercfg /hibernate off"
   ProgressSet(80,"")
   RunWait('"' & @ComSpec & '" /c ' & $CMD)
$CMD = ""
   ProgressOff()
EndFunc

Func FlushDNSCacheFunc()
   $CMD = 'ipconfig /flushdns && ' & _
        'net stop dnscache && ' & _
        'ipconfig /flushdns && ' & _
        'net start dnscache && ' & _
        'ipconfig /flushdns'
RunWait('"' & @ComSpec & '" /c ' & $CMD,"", @SW_HIDE)
$CMD = ""
EndFunc

Func DeleteCredsFunc()
   RunWait(@ComSpec & " /c " & $intracred)
	  $intracred = "CMDKEY /delete:Microsoft*"
EndFunc

Func PingFunc()
   $PingIP = InputBox($SWName,"Skriv inn IP eller adresse")
   If $PingIP > "" Then
		 $PingStatus = Ping($PingIP)
		 If $PingStatus > 0 Then
			MsgBox(64,$SWName,$PingIP&" er Online")
		 Else
			MsgBox(48,$SWName,$PingIP&" er Offline")
		 EndIf
	  Else
   EndIf
EndFunc

Func ShutdownFunc()
ShellExecute("shutdown"," -s -t 0")
EndFunc

Func RebootFunc()
ShellExecute("shutdown"," -r -t 0")
EndFunc

Func OpenChangelogFunc()
   ShellExecute(@ScriptDir&"\sys\changelog\changelog.txt")
 ;  MsgBox(0,"","")
EndFunc

Func RenewIPFunc()
   ;Edit 1.5.4  06.09.2017
RunWait (@ComSpec & " /c " & "ipconfig /release","", @SW_HIDE)
RunWait (@ComSpec & " /c " & "ipconfig /renew","", @SW_HIDE)

   #cs
   $i = "x"
   Run("cmd.exe","",@SW_SHOW)
   WinWait("C:\Windows\system32\cmd.exe")
   Send("ipconfig /release"&Chr(10)&Chr(13))
   Sleep(500)
   Send("ipconfig /renew"&Chr(10)&Chr(13))
   Send("exit"&Chr(10)&Chr(13))
   $i = WinExists("cmd.exe")
Do
   If $i = 1 Then Sleep(250)
   Until $i = 0
   GUISetState(@SW_DISABLE)
   GUISetState(@SW_ENABLE)
   Sleep(2500)
   MsgBox(64,$SWName,"IP adressen er nå byttet")
   #ce
EndFunc

Func GPResultFunc()
   $R = Random(0,$RandomTotal)
   ProgressOn($SWName,"")
   ProgressSet(25,"Høster og gjeter variabler")
$i = @MDAY&@MON&@YEAR&@HOUR&@MIN
ProgressSet(50,$RandomTXT[$R])
RunWait("gpresult /h gpresult"&$i&".html",@tempdir)
ProgressSet(75,"Lukker driten")
WinWaitClose("C:\Windows\System32\gpresult.exe")
ProgressSet(100,"Åpner gpresult")
ShellExecute(@TempDir&"\gpresult"&$i&".html")
ProgressOff()
$i = ""
EndFunc

Func RandomMessageFunc()
$i = Random(0,$RandomTotal)
MsgBox(0,$SWName,$RandomTXT[$i])
EndFunc

Func WindowsUpdateFixFunc()
Dim $WUAVersion = "WindowsUpdate Fix 2.3"
ProgressOn($WUAVersion,"")
ProgressSet(25,"Stopping Windows Update service")
RunWait (@ComSpec & " /c " & "net stop wuauserv","", @SW_HIDE)
$WUAFILE = FileExists(@ScriptDir&"\sys\sources\ResetWindowsUpdate.bat")
If $WUAFILE = 1 Then
   ProgressSet(65,"Adding .   records in registry")
   RunWait(@ScriptDir&"\sys\sources\ResetWindowsUpdate.bat","", @SW_HIDE)
Else
   ;Portable mode
   ProgressSet(65,"Sourcefile missing!"&@crlf&"To complete all parts of this script it needs to be run with all it's sources."&@CRLF&@CRLF&"Running script in Portable mode")
   Sleep(500)
EndIf
ProgressSet(75,"Creating new SoftwareDistribution folder")
DirMove("C:\Windows\SoftwareDistribution","C:\Windows\SoftwareDistributionOld")
ProgressSet(85,"Creating new logfile")
DirMove("C:\Windows\WindowsUpdate.log","C:\Windows\WindowsUpdateOld.log")
ProgressSet(95,"Starting Windows Update service")
RunWait (@ComSpec & " /c " & "net start wuauserv","", @SW_HIDE)
ProgressSet(100,"All done")
ProgressOff()
MsgBox(64,$WUAVersion,"Windows update fix is complete",5)
EndFunc

Func EmptyIVAR()
   $i = ""
EndFunc

Func KillTTool()
   ProcessClose("T-Tool.exe")
EndFunc

Func StromalternativerFunc()
   ShellExecute("powercfg.cpl")
EndFunc

Func AppWizFunc()
   ShellExecute("appwiz.cpl")
EndFunc

Func INetFunc()
   Run("inetcpl.cpl")
EndFunc

Func LusrmgrFunc()
   ;Local users and groups
   ShellExecute("lusrmgr.msc")
EndFunc

Func DeviceManager()
   ShellExecute("devmgmt.msc")
EndFunc

Func EventViewer()
   ShellExecute("eventvwr.msc")
EndFunc

Func TaskScheduler()
   ShellExecute("taskschd.msc")
EndFunc

Func DTV()
   ;$DTVFilePath = _WinAPI_GetTempFileName(@TempDir)
   ;Local $DTVFile = InetGet("https://get.teamviewer.com/torghattenqs/*.exe", $DTVFilePath, $INET_FORCERELOAD)
   ;$DTVFile = InetGet("https://get.teamviewer.com/torghattnqs",@TempDir&"\teamviewer.exe",1,0)
EndFunc

Func OpenLink()
   ShellExecute($Link)
EndFunc

Func StartOutlookInSafemode()
   ShellExecute("outlook.exe"," /safe")
EndFunc

Func ReportError()
   ShellExecute("mailto:"&$MailAddress&"?subject=Feil med "&$SWName&" "&$SWVersion)
EndFunc

Func BookmarkCreate()
If $ShortcutFile > "" Then
FileCopy(@ScriptDir&"\sys\sources\shortcuts\"&$ShortcutFile,@DesktopDir&"\"&$ShortcutFile,0)
Else
   MsgBox(48,$SWName,"No shortcut found")
EndIf
EndFunc

Func CredentialManager()
   Run("control.exe /name Microsoft.CredentialManager")
EndFunc

Func ResMon()
   ShellExecute("resmon")
EndFunc

Func TaskManager()
   ShellExecute("taskmgr")
EndFunc

Func RepairOffice365()
   ;RepairOffice365 Version 1.1
   MsgBox(64,$SWName,"This function is merely tested once")
   If FileExists("C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe") Then
	  ShellExecute("C:\Program Files\Microsoft Office 15\ClientX64\OfficeClickToRun.exe","scenario=Repair platform=x86 culture=no-nb RepairType=FullRepair DisplayLevel=True")
   Else
	  MsgBox(16,$SWName,"Ojsann! Her gikk noe galt."&@CRLF&"Automatisk reparasjon fungerer ikke på denne enheten."&@CRLF&"Åpner Programmer og Funksjoner...")
	  ShellExecute("appwiz.cpl")
   EndIf
EndFunc

Func PortableCheck()
   If FileExists(@scriptdir&"\sys") Then
	  $Portable = ""
   Else
	  $Portable = " - (Portable mode)"
   EndIf
EndFunc

Func DevicesAndPrinters()
   Run("control printers")
EndFunc

Func RegEdit()
   Run("regedit")
EndFunc

Func WinUpdate()
   Run("control /name Microsoft.WindowsUpdate")
EndFunc

Func WinFirewall()
   Run ("control firewall.cpl")
EndFunc

Func GetNSIP()
	  ;edit 1.5.4.5
		 runWait(@comSpec & ' /c nslookup '&$Hostname&' > ' & @tempDir & '\nslookup.txt"',@ScriptDir,@SW_HIDE)
		 $i = 0
	  Do
		 $i = $i + 1
			$NSIP = FileReadLine(@tempDir & "\nslookup.txt",$i)
			$NSIP = StringLeft($NSIP,9)
			If $NSIP = "Addresses" Then $i = $i + 100
	  Until $i > 100
		 $i = $i - 100
		 $NSIP = FileReadLine(@tempDir & "\nslookup.txt",$i)
			$NSIP = StringTrimLeft($NSIP,10)
			   $NSIP = StringStripWS($NSIP,$STR_STRIPALL)
	 ; MsgBox(0,$SWName,$NSIP,5)

	  If $NSIP = "" Then $NSIP = @IPAddress1
		 If StringIsAlpha($NSIP) = 0 Then $NSIP = @IPAddress1
		 GUICtrlSetData($IPAddress,"IP Adresse: "&$NSIP)
   EndFunc

Func BetaMode()
   $BetaCheck = MsgBox(52+256,$SWName,"Denne funksjonen kjøres i beta-modus."&@CRLF&"Er du helt sikker på at du vil bruke denne funksjonen?"&@CRLF&@CRLF&"Funksjoner i beta-modus kan potensiellt skade maskinen om man ikke vet hva man gjør eller hvordan funksjonene virker.")

   If $BetaCheck = 6 then $BetaCheck = "Yes"
	  If $BetaCheck = 7 then $BetaCheck = "No"
   ;MsgBox(64,$SWName,"Du trykket akkurat på: "&$BetaCheck)
EndFunc

Func Success()
   MsgBox(0,$SWName,"Success!")
EndFunc

Func SetSafemode()

   If FileExists(@Scriptdir&"\sys\sources\safemode.txt") Then
	  $Safemode = FileReadLine(@Scriptdir&"\sys\sources\safemode.txt",2)
	  $Safemode = StringRight($Safemode,1)
	  ;MsgBox(0,$SWName,$Safemode) ; gir output om hva verdien i safemode er
	  If $Safemode = "" Then _FileWriteToLine(@Scriptdir&"\sys\sources\safemode.txt",2,"safemode: 1",True)

		If $Safemode = 1 Then
			GUICtrlSetState($MenuSafeMode,$GUI_CHECKED)
			;MsgBox(0,"","Setting state")
		 Else
			GUICtrlSetState($MenuSafemode,$GUI_UNCHECKED)
		 EndIf

   Else
	$Safemode = 1
	  If $Safemode = 1 Then
		 GUICtrlSetState($MenuSafeMode,$GUI_CHECKED)
	  Else
		 GUICtrlSetState($MenuSafemode,$GUI_UNCHECKED)
	  EndIf
   EndIf


EndFunc

Func ChangeSafemode()
   If FileExists(@Scriptdir&"\sys\sources\safemode.txt") Then
	  $Safemode = FileReadLine(@Scriptdir&"\sys\sources\safemode.txt",2)
	  $Safemode = StringRight($Safemode,1)

		 If $Safemode = 1 Then
			$Safemode = 0
			_FileWriteToLine(@Scriptdir&"\sys\sources\safemode.txt",2,"safemode: 0",True)
			GUICtrlSetState($MenuSafeMode,$GUI_UNCHECKED)
		 Else
			$Safemode = 1
			_FileWriteToLine(@Scriptdir&"\sys\sources\safemode.txt",2,"safemode: 1",True)
			GUICtrlSetState($MenuSafemode,$GUI_CHECKED)
		 EndIf

	  ;MsgBox(0,$SWName,$Safemode)
   Else
	  ;Hvis safemode.txt ikke finnes
	  If $Safemode = 1 Then
			$Safemode = 0
			GUICtrlSetState($MenuSafeMode,$GUI_UNCHECKED)
		 Else
			$Safemode = 1
			GUICtrlSetState($MenuSafemode,$GUI_CHECKED)
		 EndIf
   EndIf
EndFunc

Func SafemodeApps()
   If $Safemode = 1 Then
		 GUICtrlSetState($MenuDeleteCreds,$GUI_DISABLE)
			GUICtrlSetState($BTN_DelTemp,$GUI_DISABLE)


   Else
		 GUICtrlSetState($MenuDeleteCreds,$GUI_ENABLE)
			GUICtrlSetState($BTN_DelTemp,$GUI_ENABLE)

   EndIf

EndFunc

Func ResyncTime()
runWait(@comSpec & ' /c w32tm /resync',@ScriptDir,@SW_HIDE)
EndFunc

Func GodMode()
   ;Windows XP miljø går rett i BOD om man kjører GodMode
   If $WinVersjon = "Windows XP" Then
	  MsgBox(16,$SWName,"GodMode fungerer ikke i Windows XP miljø.")
   Else
	  ShellExecute("shell:::{26EE0668-A00A-44D7-9371-BEB064C98683}\0\::{ed7ba470-8e54-465e-825c-99712043e01c}")
   EndIf
EndFunc


Func OpprettSak()
   ShellExecute("mailto:support@torghatten.no")
EndFunc

Func EventLog()
    Local $hEventLog, $aEvent

    ; Create GUI
    GUICreate("EventLog", 400, 300)
    $g_idMemo = GUICtrlCreateEdit("", 2, 2, 396, 300, 0)
    GUICtrlSetFont($g_idMemo, 9, 400, 0, "Courier New")
    GUISetState(@SW_SHOW)

    ; Read most current event record
    ;$hEventLog = _EventLog__Open("", "Application")
    ;$aEvent = _EventLog__Read($hEventLog, True, False) ; read last event
     $hEventLog = _EventLog__Open("", "System")
     $aEvent = _EventLog__Read($hEventLog)
    ; $aEvent = _EventLog__Read($hEventLog, True, False)
    MemoWrite("Result ............: " & $aEvent[0])
    MemoWrite("Record number .....: " & $aEvent[1])
    MemoWrite("Submitted .........: " & $aEvent[2] & " " & $aEvent[3])
    MemoWrite("Generated .........: " & $aEvent[4] & " " & $aEvent[5])
    MemoWrite("Event ID ..........: " & $aEvent[6])
    MemoWrite("Type ..............: " & $aEvent[8])
    MemoWrite("Category ..........: " & $aEvent[9])
    MemoWrite("Source ............: " & $aEvent[10])
    MemoWrite("Computer ..........: " & $aEvent[11])
    MemoWrite("Username ..........: " & $aEvent[12])
    MemoWrite("Description .......: " & $aEvent[13])
    _EventLog__Close($hEventLog)

    ; Loop until user input.
    Do
    Until GUIGetMsg() = $GUI_EVENT_CLOSE
 EndFunc

 Func MemoWrite($sMessage)
    GUICtrlSetData($g_idMemo, $sMessage & @CRLF, 1)
 EndFunc

 Func DelTemp()
	ProgressOn($SWName,"","Sletter %temp%")
	runWait(@comSpec & ' /c DEL /F /S /Q %TEMP%',@ScriptDir,@SW_HIDE)
	  ProgressSet(25,"Sletter .tmp filer")
	  runWait(@comSpec & ' /c DEL C:\*.tmp /S /Q',@ScriptDir,@SW_HIDE)
		 ProgressSet(50,"Sletter området %tmp%")
		 runWait(@comSpec & ' /c DEL %tmp% /S /Q',@ScriptDir,@SW_HIDE)
			ProgressSet(75,"Sletter området %temp%")
			runWait(@comSpec & ' /c DEL %temp% /S /Q',@ScriptDir,@SW_HIDE)
   ProgressOff()
EndFunc

Func WideSN()
;*5 for antall pixler i en karakter, 10 er maxgrensa for antall tegn
;$SN = $SN & " JODANÅGÅRDETFINT"
If StringLen($SN) > 12 Then
$GUIWidth = $GUIWidth + ((StringLen($SN) - 10) * 6)
$SNWidth = $SNWidth + ((StringLen($SN) - 10) * 6)

Else
$GUIWidth = $GUIWidth
EndIf

EndFunc

Func OpenHosts()
Run("notepad.exe "&$System32&"Drivers\etc\hosts")
EndFunc

Func CheckDeviceDomainStatus()
   $DeviceDomainStatus = @LogonDomain
   If $DeviceDomainStatus = "TORGHATTEN" Then
	  $DeviceDomainStatus = "Maskina er i Torghatten-domenet"
   Else
	  $DeviceDomainStatus = "Maskina er ikke i Torghatten-domenet"
   EndIf
   ;MsgBox(0,"",$DeviceDomainStatus)
EndFunc

Func KillOneDriveEXE()
  If ProcessExists("OneDrive.exe") Then ProcessClose("OneDrive.exe")
  EndFunc

Func GetComputerModel()
      runWait(@comSpec & ' /c wmic csproduct get vendor> "' & @tempDir & '\csproduct.txt"',@ScriptDir,@SW_HIDE)
	  runWait(@comSpec & ' /c wmic csproduct get version >> "' & @tempDir & '\csproduct.txt"',@ScriptDir,@SW_HIDE)
	  runWait(@comSpec & ' /c wmic computersystem get model > "' & @tempDir & '\computersystemMM.txt"',@ScriptDir,@SW_HIDE)
   $Vendor = FileReadLine(@tempDir&"\csproduct.txt",2)
   $Vendor = StringStripWS($Vendor,$STR_STRIPALL)
   $VendorVersion = FileReadLine(@tempDir&"\csproduct.txt",4)
   $VendorVersion = StringStripWS($VendorVersion,$STR_STRIPALL)
   $Vendor = StringReplace($Vendor,"FUJITSU","HP Rebrand")
   $ComputerModelnumber = FileReadLine(@tempDir&"\computersystemMM.txt",2)
   $ComputerModelnumber = StringStripWS($ComputerModelnumber,$STR_STRIPALL)

   If $ComputerModelExtended = 1 Then
   MsgBox(64,$SWName,"Vendor: "&$Vendor&@CRLF&"Version: "&$VendorVersion&@CRLF&"Modelnumber: "&$ComputerModelnumber)
   $ComputerModelExtended = 0
   EndIf

EndFunc

Func ShowFileExt()

   $FileExtStatus = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt")

If $FileExtStatus = 1 Then
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", "REG_DWORD", 0)
Else
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",  "HideFileExt", "REG_DWORD", 1)
EndIf

EndFunc

Func ShowHideExt()
   ;Denne finner bare ut hva som skal stå i knappen
   $FileExtStatus = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt")
   If $FileExtStatus = 1 Then
	  $FileExtStatusTag = "Vis"
   Else
	  $FileExtStatusTag = "Skjul"
   EndIf

If $GUIStatus = 1 Then GUICtrlSetData($BTN_FileExt,$FileExtStatusTag&" filtyper")
EndFunc

Func ShowHiddenFolder()

   $HiddenFolderStatus = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden")

If $HiddenFolderStatus = 1 Then
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden", "REG_DWORD", 0)
Else
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced",  "Hidden", "REG_DWORD", 1)
EndIf

EndFunc

Func ShowHideFolders()
   ;Denne finner bare ut hva som skal stå i knappen
   $HiddenFolderStatus = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "Hidden")
   If $HiddenFolderStatus = 1 Then
	  $HiddenFolderStatusTag = "Skjul"&@CRLF
   Else
	  $HiddenFolderStatusTag = "Vis"&@CRLF
   EndIf

If $GUIStatus = 1 Then GUICtrlSetData($BTN_HiddenFolder,$HiddenFolderStatusTag&" skjulte mapper")
EndFunc

Func GetDNS()
runWait(@comSpec & ' /c ipconfig /all > ' & @TempDir & '\ipconfig.txt"',@ScriptDir,@SW_HIDE)
;14 chrs = DNS Servers
$i = 0
Do
   $i = $i + 1
   $DNSReadLine = FileReadLine(@TempDir&"\ipconfig.txt",$i)
   If StringLeft($DNSReadLine,14) = "   DNS Servers" then $i = $i + 100
   Until $i > 100
$DNSReadLine = StringRight($DNSReadLine,StringLen($DNSReadLine)-StringInStr($DNSReadLine,":"))
StringStripWS($DNSReadLine,$STR_STRIPALL)

If $DNSReadLine > "" Then

$DNSPrim = $DNSReadLine
   $DNSPrim = StringStripWS($DNSPrim,$STR_STRIPALL)
	  $DNSSec = FileReadLine(@TempDir&"\ipconfig.txt",$i-100+1)
		 $DNSSec = StringStripWS($DNSSec,$STR_STRIPALL)

MsgBox(0,$SWName,"Primary DNS: " & $DNSPrim & @CRLF & "Secondary DNS: " & $DNSSec)

Else
   MsgBox(48,$SWName,"Finner ikke DNS")
EndIf

EndFunc
StringIsAlNum(
Func FindUser() ;v1.1
   ;Her sjekker skriptet om brukernavn matcher brukernavnet i scriptdirpath
   $i = 0
   $DirUser = @ScriptDir
   $DirUser = StringTrimLeft($DirUser,3)
   If StringLeft($DirUser,5) = "users" Then
	  $DirUser = StringTrimLeft($DirUser,6)
	  $i = StringInStr($DirUser,"\")
	  $DirUser = StringLeft($DirUser,$i-1)

   If $DirUser = @UserName&"d" Then
	  ;UsernamesMatch
	  $DirUser = @UserName
   Else
;	  $i = MsgBox(0,"","Usernames do not match, using path username: "&$DirUser&@CRLF&"Other username is: "&@UserName)

	  GUICtrlSetData($UsernameLabel,"Brukernavn: "&$DirUser)
   EndIf
$i = 0
   EndIf
EndFunc

Func RunCommand()
$RunInput = InputBox($SWName,"skriv inn kommando")
If @error = 1 Then
   ""
Else
ShellExecute($RunInput)
EndIf

EndFunc

Func CMD()
   RunWait("cmd.exe")
EndFunc