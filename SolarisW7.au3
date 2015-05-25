#NoTrayIcon
#include <_XMLDomWrapper.au3>
#include <file.au3>
#include <Array.au3>
#Include <GuiToolBar.au3>
#include <Process.au3>
#include <Date.au3>
#include <WindowsConstants.au3>
#Include <StaticConstants.au3>
#include <GuiToolTip.au3> 

Global $Exten=0, $AvayaId=0, $hWnd, $w_width=750, $w_height=480

_SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "',DEFAULT,13,DEFAULT,DEFAULT," & "'start solaris'", "event.id")

DirRemove(@AppDataDir & "\Avaya", 1)
DirCreate(@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default")
DirCreate("C:\Program Files (x86)\Avaya\Avaya Aura CC Elite Multichannel\Desktop\CC Elite Multichannel Desktop")

;If Not FileExists (@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Settings.xml") Then
;     DirCreate (@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default")
;	Else 
;	 FileSetAttrib(@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Settings.xml", "-RS") 
;	 FileSetAttrib(@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\ScreenPops.xml", "-RS") 
;EndIf

FileInstall("Settings.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Settings.xml", 1)
FileInstall("ScreenPops.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\ScreenPops.xml", 1)
FileInstall("Preferences.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Preferences.xml", 1)
FileInstall("AuxReasonCodes.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\AuxReasonCodes.xml", 1)
FileInstall("Abbreviations.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Abbreviations.xml", 1)
FileInstall("AudioGreetings.xml", @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\AudioGreetings.xml", 1)
FileInstall("ASGUIHost.ini", "C:\Program Files (x86)\Avaya\Avaya Aura CC Elite Multichannel\Desktop\CC Elite Multichannel Desktop\ASGUIHost.ini", 1)

$sXMLFile = @AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Settings.xml"
$name_space = '"http://avaya.com/OneXAgent/Settings"'

$Open = _XMLFileOpen($sXMLFile, 'xmlns=' & $name_space)
$objDoc.setProperty("SelectionNamespaces", 'xmlns:MyNS=' & $name_space) ; даем имя MyNS дефолтному namespace для его последующего использования

$sINIFile = "C:\Program Files (x86)\Avaya\Avaya Aura CC Elite Multichannel\Desktop\CC Elite Multichannel Desktop\ASGUIHost.ini"

;MsgBox(64, "Test", "Childs: " & $aChilds)
;MsgBox(64, "Test", "Code: " & $Open & " Extended: " & @Extended)

;$Station =_XMLGetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Telephony/MyNS:User', 'Station')
;$Agent = _XMLGetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Agent', 'Login')
;MsgBox(64, "Test", "Station: " & $Station & " Agent: " & $Agent & @CRLF & "Error: " & @Error)

;$AvayaId = 6477

$aProcessList = ProcessList("SolarisW7.exe")
If $aProcessList[0][0] > 1 Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",13,DEFAULT,DEFAULT," & "'solaris double launch detected'", "event.id")
   _WriteLog("Double launch detected. Exiting...")
   MsgBox(64, 'SOLARIS', 'Копия программы SOLARIS уже запущена. Выйдите из программы или дождитесь завершения её работы.', 4)
   _Exit()
EndIf

If Not _CalcExten() Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "',DEFAULT,13,DEFAULT,DEFAULT," & "'unable to calculate extension'", "event.id")
   MsgBox(64, 'SOLARIS', 'Не удалось определить Extension. Проверьте имя компьютера.', 4)
   _Exit()
EndIf

If ProcessExists("OneXAgentUI.exe") Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",13,DEFAULT,DEFAULT," & "'one-x-agent already runned'", "event.id")
   _WriteLog("OneXAgentUI.exe already runned. Exiting...")
   MsgBox(64, 'SOLARIS', 'Обнаружен активный One-X Agent. Запуск Solaris невозможен. Приложение будет завершено.', 4)
   $hWnd = WinWait("Avaya one-X Agent", "", 5)
   _CloseAgent()
   _Exit()
EndIf

If FileGetSize(@TempDir & "\solaris.log") > 10485760 Then
   FileDelete(@TempDir & "\solaris.log")
EndIf

If FileExists (@TempDir & "\busy.flag") Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",12,DEFAULT,DEFAULT," & "'termination detected'", "event.id")
   _WriteLog('[' & @ComputerName & '\' & @UserName & "] Abnormal program termination on previous run detected")
   MsgBox(64, 'SOLARIS', "Предыдущее завершение программы было некорректным!", 3)
Else
   FileClose (FileOpen (@TempDir & "\busy.flag", 2))
EndIf

If FileExists(@TempDir & "\agent.id") Then
   FileDelete(@TempDir & "\agent.id")
EndIf

_WriteLog('[' & @ComputerName & '\' & @UserName & "] -=-=-=-=-=-=-=-Running  Solaris-=-=-=-=-=-=-=-")
_SQLExec("exec SOLARIS_2.dbo.GET_AVAYA_ID_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten, "agent.id")

$file = FileOpen(@TempDir & "\agent.id", 0)
$AvayaID = StringStripWS (FileReadLine($file, 3), 1)

;MsgBox(64, 'SOLARIS', $AvayaID)

If StringLen($AvayaID) <> 4 or StringLeft($AvayaID, 1) <> '6' Then
   _WriteLog("Unable to get Agent ID")
   MsgBox(64, 'SOLARIS', "Не удалось получить идентификатор агента")
   ;MsgBox(64, 'SOLARIS', $AvayaID)
   _Exit()
EndIf

_WriteLog("Issued Agent ID: " & $AvayaID)

_XMLSetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Telephony/MyNS:User', 'Station', $Exten)
_XMLSetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Agent', 'Login', $AvayaId)

IniWrite($sINIFile, "Telephony", "Station DN", " " & $Exten)
IniWrite($sINIFile, "User", "Agent ID", " " & $AvayaId)

;FileSetAttrib(@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\Settings.xml", "+RS") 
;FileSetAttrib(@AppDataDir & "\Avaya\one-X Agent\2.5\Profiles\default\ScreenPops.xml", "+RS") 

;MsgBox(64, "Test", "RO")

Run("C:\Program Files (x86)\Avaya\Avaya one-X Agent\OneXAgentUI.exe", "C:\Program Files (x86)\Avaya\Avaya one-X Agent")

SplashTextOn("SOLARIS", "Подготовка рабочего места оператора, ожидайте..." & @CRLF & @CRLF & "Просьба не пользоваться мышкой и клавиатурой пока не погаснет это сообщение", $w_width, $w_height, (@DesktopWidth-$w_width)/2, (@DesktopHeight-$w_height)/2, 4, "", 24)

$hWnd = WinWait("Добро пожаловать! - Avaya one-X Agent", "", 30)

If Not $hWnd Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",14,DEFAULT,DEFAULT," & "'failed to start one-x-agent'", "event.id")
   SplashOff()
   MsgBox(64, 'SOLARIS', 'Не удалось запустить Avaya One-X Agent', 4)
   _Exit()
   Else
      WinActivate("Добро пожаловать! - Avaya one-X Agent")
	  Send("{ENTER}")
EndIf

$hWnd = WinWait("Вход в систему в качестве станции - Avaya one-X Agent", "", 15)

If Not $hWnd Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",2,DEFAULT,DEFAULT," & "'failed to register the station'", "event.id")
   SplashOff()
   MsgBox(64, 'SOLARIS', 'Не удалось зарегистрировать станцию', 4)
   _Exit()
   Else
      WinActivate("Вход в систему в качестве станции - Avaya one-X Agent")
	  Send("0000")
	  Send("{ENTER}")
EndIf

$hWnd = WinWait("Вход в систему в качестве оператора - Avaya one-X Agent", "", 15)

If Not $hWnd Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",3,DEFAULT,DEFAULT," & "'failed to register the operator'", "event.id")
   SplashOff()
   MsgBox(64, 'SOLARIS', 'Не удалось зарегистрировать оператора', 4)
   _Exit()
   Else
      WinActivate("Вход в систему в качестве оператора - Avaya one-X Agent")
	  Send($AvayaId)
	  Send("{ENTER}")
EndIf

SplashOff()

$hWnd = WinWait("Avaya one-X Agent", "", 5)

While 1
   If WinExists($hWnd) = 0 Then
	  _Exit()
   EndIf
   Sleep(20)   
WEnd

;============================================================================================
Func _Exit()
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",4,DEFAULT,DEFAULT," & "'exit'", "event.id")
   _WriteLog('[' & @ComputerName & '\' & @UserName & "] -=-=-=-=-=-=-=-=-=-=-EXIT-=-=-=-=-=-=-=-=-=-=-" & @CRLF)
   FileDelete (@TempDir & "\busy.flag")
;   MsgBox(64, "SOLARIS", "Завершение работы", 3)
   Exit
EndFunc 

Func _CalcExten()
   $IPaddr1 = @IPAddress1
   $IPaddr2 = @IPAddress2
   $IPaddr3 = @IPAddress3
   $IPaddr4 = @IPAddress4
   $IPaddr = ""
   $OctetsIP = ""
   $CompN = ""
   
   $CName = @ComputerName
   $Octet = StringMid($CName, 3, 3)
   
   If $IPaddr4 = "0.0.0.0" Then
	If $IPaddr3 = "0.0.0.0" Then
		If $IPaddr2 = "0.0.0.0" Then
			If ($IPaddr1 = "0.0.0.0") Or ($IPaddr1 = "127.0.0.1") Then
				MsgBox(64, "SOLARIS", "Ошибка IP-адреса. Проверьте настройки сетевой карты")
				_Exit()
			Else 
				$IPaddr = $IPaddr1
			EndIf
		Else 
			$IPaddr = $IPaddr2
		EndIf
	Else
		$IPaddr = $IPaddr3
	EndIf
   Else
	$IPaddr = $IPaddr4
   EndIf
   
   ;MsgBox(64, "SOLARIS", $IPaddr)
   $OctetsIP = StringSplit($IPaddr, ".")
   ;_ArrayDisplay($OctetsIP, 'Октеты')

   $CompN = "FL" & StringFormat("%03i", $OctetsIP[3]) & StringFormat("%03i", $OctetsIP[4])

   ;MsgBox(64, "SOLARIS", $CompN)
   
   If $CompN <> $CName Then
	MsgBox(64, "SOLARIS", "Имя компьютера не соответствует IP-адресу. Обратитесь к специалистам технической поддержки!")
    _Exit()
   EndIf
   
   Select 
   Case $Octet >= 100 AND $Octet <= 103 
	  $Numb = (StringMid($CName, 5, 1)*300 + StringMid($CName, 6, 3))
	  $Numb = StringFormat("%03i", $Numb)
	  $Exten = "5" & $Numb
	  Return  $Exten
   ; MsgBox(64, "SOLARIS", $Octet)
   Case $Octet = "020"
	  $Exten = StringMid($CName, 6, 3) + 1200
	  Return  $Exten
   Case Else
	  Return 0
   EndSelect
EndFunc

Func _WriteLog($Msg)
   $file = FileOpen (@TempDir & "\solaris.log", 1)
   FileWriteLine($file, @HOUR & ':' & @MIN & ':' & @SEC & ' ' & @MDAY & '/' & @MON & '/' & @YEAR & " " & $Msg & @CRLF)
   FileClose($file)
EndFunc

Func _SQLExec($CmdString, $OutFName)
   $file = FileOpen(@TempDir & "\template.sql", 2)
   FileWriteLine($file, "SET NOCOUNT ON")
   FileWriteLine($file, "GO")   
   FileWrite($file, $CmdString)
   FileClose($file)

;MsgBox(64, 'SOLARIS', $FileData, 4)

   RunWait("C:\Program Files\Microsoft SQL Server\100\Tools\Binn\sqlcmd.exe -S mars -U task_executor -P Up1864 -i " & @TempDir & "\template.sql -o " & @TempDir & "\" & $OutFName, "C:\Program Files\Microsoft SQL Server\100\Tools\Binn", @SW_HIDE)
EndFunc

Func _CloseAgent()
   BlockInput (1)
   WinActivate($hWnd)
   Sleep(500)
   WinClose($hWnd)
   WinWaitClose($hWnd, 7)
   If ProcessExists("OneXAgentUI.exe") Then
      _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",13,DEFAULT,DEFAULT," & "'killing process OneXAgentUI.exe'", "event.id")
	  _WriteLog("Killing process OneXAgentUI.exe")
	  ProcessClose("OneXAgentUI.exe")
   EndIf
   BlockInput (0)
EndFunc
