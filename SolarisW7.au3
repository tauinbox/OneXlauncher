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
#Include <GUIConstantsEx.au3>
#Include <ButtonConstants.au3>
#include <MsgBoxConstants.au3>

Global $Exten=0, $AvayaId=0, $hWnd, $hWndPop, $w_width=750, $w_height=480, $OpName="", $ChatState=0, $Answer, $OpMode=0, $ChatGreeting=0, $Pos, $X, $Y, $IsDataCorrect=0
Global $Greeting, $Ok=0

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

_SQLExec("exec SOLARIS_2.dbo.GET_OPERATOR_FIO " & "'" & @UserName & "'", "opname.id")

$file = FileOpen(@TempDir & "\opname.id", 0)
$OpName = StringStripWS (FileReadLine($file, 3), 1)

_WriteLog("Issued Agent ID: " & $AvayaID)
_WriteLog("Operator Name: " & $OpName)
_SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",13,DEFAULT,DEFAULT," & "'user: " & $OpName & "'", "event.id")

_XMLSetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Telephony/MyNS:User', 'Station', $Exten)
_XMLSetAttrib('/MyNS:Settings/MyNS:Login/MyNS:Agent', 'Login', $AvayaId)

IniWrite($sINIFile, "Telephony", "Station DN", " " & $Exten)
IniWrite($sINIFile, "User", "Agent ID", " " & $AvayaId)
;IniWrite($sINIFile, "Simple Messaging", "Agent Specific Welcome Message", " " & "Здравствуйте, меня зовут " & StringStripWS ($OpName, 2) & ", специалист технической поддержки абонентов " & '"Триколор ТВ"')

While NOT $IsDataCorrect
	$Opmode = 0
	$Greeting = " Здравствуйте, меня зовут " & StringStripWS ($OpName, 2) & ", специалист технической поддержки"
	_SQLExec("exec SOLARIS_2.dbo.GET_WEBCHAT_STATE_v2 " & "'" & @UserName & "'", "chatstate.id")

	$file = FileOpen(@TempDir & "\chatstate.id", 0)
	$ChatState = StringStripWS (FileReadLine($file, 3), 8)
	
	Choice("Выбор режима работы", "Выберите требуемый режим работы!")
	
	$ChatState = Number($ChatState)

	If $Answer = 2 Then
		If $ChatState <> 0 Then
			_XMLSetAttrib('/MyNS:Settings/MyNS:WorkHandling/MyNS:Accept', 'AutoAccept', 'false')
		EndIf
		If $ChatState = 0 Then
			If _IsItOk("Кажется, что-то пошло не так", "Мы обнаружили расхождение выбранного Вами режима работы с подключенными Вам каналами обслуживания. Вы выбрали режим работы с чатами, хотя нам не удалось обнаружить ни одной активной линии чатов, назначенной для Вас." & @CR & @CR & "К сожалению, ЭВМ тоже иногда ошибается, поэтому подскажите, как следует поступить в данной ситуации?") Then
				_XMLSetAttrib('/MyNS:Settings/MyNS:WorkHandling/MyNS:Accept', 'AutoAccept', 'false')
				$IsDataCorrect = 1
			EndIf
		ElseIf $ChatState = 1 Then
			If $ChatGreeting = 0 Then
				$Greeting = " Здравствуйте, меня зовут " & StringStripWS ($OpName, 2) & ", специалист технической поддержки абонентов " & '"Триколор ТВ".'
				$IsDataCorrect = 1
			EndIf
			If $ChatGreeting <> 0 Then
				If _IsItOk("Кажется, что-то пошло не так", "Вам не назначена линия обработки чатов по каналу GS Gamekit, хотя Вы пытаетесь зарегистрироваться в системе с указанием данной опции." & @CR & @CR & "К сожалению, ЭВМ тоже иногда ошибается, поэтому подскажите, как следует поступить в данной ситуации?") Then
					$IsDataCorrect = 1
				EndIf
			EndIf
		ElseIf $ChatState = 10 Then
			If $ChatGreeting = 1 Then
				$Greeting = " Здравствуйте. Служба поддержки GS Gamekit. Меня зовут " & StringStripWS ($OpName, 2)
				$IsDataCorrect = 1
			EndIf
			If $ChatGreeting <> 1 Then
				If _IsItOk("Кажется, что-то пошло не так", "Для Вас назначена линия обработки чатов по каналу GS Gamekit, хотя Вы пытаетесь зарегистрироваться в системе без указания данной опции." & @CR & @CR & "К сожалению, ЭВМ тоже иногда ошибается, поэтому подскажите, как следует поступить в данной ситуации?") Then
					$IsDataCorrect = 1
				EndIf
			EndIf
		EndIf
		;MsgBox(4096, "Error", _XMLError ())
	ElseIf $Answer = 1 Then
		_XMLSetAttrib('/MyNS:Settings/MyNS:WorkHandling/MyNS:Accept', 'AutoAccept', 'true')
		If $ChatState <> 0 Then
			If _IsItOk("Кажется, что-то пошло не так", "Для Вас назначена линия обработки чатов, хотя Вы пытаетесь зарегистрироваться в системе без возможности обработки чатов." & @CR & @CR & "К сожалению, ЭВМ тоже иногда ошибается, поэтому подскажите, как следует поступить в данной ситуации?") Then
				;MsgBox(64, "Активен канал обработки чатов", "Для Вас назначена линия обработки чатов, вход в систему возможен только в режиме обработки онлайн чатов")
				$IsDataCorrect = 1
			EndIf
		ElseIf $ChatState = 0 Then
			$IsDataCorrect = 1
		EndIf
	Endif

	If $Opmode = 1 Then
		IniWrite($sINIFile, "Rules", "Rule5", " " & "When LinkedRule Always Do SendKeys CC Elite*,{F10} Then Stop Else Stop")
	Else
		IniWrite($sINIFile, "Rules", "Rule5", " " & "When LinkedRule Always Do SendKeys CC Elite*,{F11} Then Stop Else Stop")
	EndIf
	
	IniWrite($sINIFile, "Simple Messaging", "Agent Specific Welcome Message", $Greeting)
WEnd

;MsgBox(64, "Chatstate", $ChatState)

;_Exit()

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
	  BlockInput(1)
      WinActivate("Добро пожаловать! - Avaya one-X Agent")
	  Send("{ENTER}")
	  BlockInput(0)
EndIf

$hWnd = WinWait("Вход в систему в качестве станции - Avaya one-X Agent", "", 15)

If Not $hWnd Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",2,DEFAULT,DEFAULT," & "'failed to register the station'", "event.id")
   SplashOff()
   MsgBox(64, 'SOLARIS', 'Не удалось зарегистрировать станцию', 4)
   _Exit()
   Else
      BlockInput(1)
      WinActivate("Вход в систему в качестве станции - Avaya one-X Agent")
	  Send("0000")
	  Send("{ENTER}")
	  BlockInput(0)
EndIf

$hWnd = WinWait("Вход в систему в качестве оператора - Avaya one-X Agent", "", 15)

If Not $hWnd Then
   _SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",3,DEFAULT,DEFAULT," & "'failed to register the operator'", "event.id")
   SplashOff()
   MsgBox(64, 'SOLARIS', 'Не удалось зарегистрировать оператора', 4)
   _Exit()
   Else
      BlockInput(1)
      WinActivate("Вход в систему в качестве оператора - Avaya one-X Agent")
	  Send($AvayaId)
	  Send("{ENTER}")
	  BlockInput(0)
EndIf

$hWnd = WinWait("Avaya one-X Agent", "", 5)

$hWndPop = WinWait("Информация - Avaya one-X Agent", "", 20)
WinWaitClose($hWndPop, "", 40)

Sleep(5000)

;MsgBox(64, "test", "Done")
;ControlClick($hWnd, "", "", "left", 1, 480, 80)

$Pos = WinGetPos($hWnd)
$X = 480
$Y = 80

SplashOff()

BlockInput(1)
WinActivate($hWnd)
MouseClick("left", $Pos[0] + $X, $Pos[1] + $Y)
BlockInput(0)

Sleep(2000)

If $Answer = 2 Then
	_WriteLog("starting EMC...")
	_SQLExec("exec SOLARIS_2.dbo.EVENT_LOG_INSERT_v2 " & "'" & @UserName & "','" & @ComputerName & "'," & $Exten & ",13,DEFAULT,DEFAULT," & "'starting EMC...'", "event.id")
	Run("C:\Program Files (x86)\Avaya\Avaya Aura CC Elite Multichannel\Desktop\CC Elite Multichannel Desktop\ASGUIHost.exe", "C:\Program Files (x86)\Avaya\Avaya Aura CC Elite Multichannel\Desktop\CC Elite Multichannel Desktop", @SW_MAXIMIZE)
EndIf

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
   Case $Octet = "021"
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

Func Choice($Title, $Note)
	Local $Btn1ID, $Btn2ID, $ChkBox_is_Master, $ChkBox_GameConsole, $msg

	$dWnd = GUICreate($Title, 520, 140)

	GUICtrlCreateIcon('user32.dll', 102, 10, 10)
	GUICtrlCreateLabel($Note, 60, 20)
	$Btn1ID = GUICtrlCreateButton("Я принимаю только голосовые вызовы", 10, 60)
	$Btn2ID = GUICtrlCreateButton("Я дополнительно обрабатываю онлайн чаты", 250, 60)
	$ChkBox_GameConsole = GUICtrlCreateCheckbox('а еще я обслуживаю чаты "игровой консоли"', 250, 90)
	$ChkBox_is_Master = GUICtrlCreateCheckbox("режим мастера (конечно же, я знаю что это)", 250, 110)

	GUISetState() ; display the GUI

	Do
		$msg = GUIGetMsg()

		Select
			Case $msg = $Btn1ID
				$Answer = 1
				If _IsChecked($ChkBox_is_Master) Then
					$OpMode = 1
				EndIf
				Exitloop
			Case $msg = $Btn2ID
				$Answer = 2
				If _IsChecked($ChkBox_is_Master) Then
					$OpMode = 1
				EndIf
				If _IsChecked($ChkBox_GameConsole) Then
					$ChatGreeting = 1
				EndIf				
				Exitloop			
			Case $msg = $GUI_EVENT_CLOSE
				$Answer = 1
				_Exit()
				Exitloop
		EndSelect
	Until $msg = $GUI_EVENT_CLOSE
	GUIDelete($dWnd)
EndFunc 

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc   ;==>_IsChecked

Func _IsItOk($Title, $Note)
	Local $Btn1ID, $Btn2ID

	$dWnd = GUICreate($Title, 520, 140)

	GUICtrlCreateIcon('user32.dll', 102, 10, 10)
	GUICtrlCreateLabel($Note, 60, 20, 450, 80)
	$Btn1ID = GUICtrlCreateButton("Да, всё в порядке, продолжаем", 10, 110)
	$Btn2ID = GUICtrlCreateButton("Я, кажется, что-то напутал, попробую заново", 250, 110)

	GUISetState() ; display the GUI

	Do
		$msg = GUIGetMsg()

		Select
			Case $msg = $Btn1ID
				$Ok = 1
				Exitloop
			Case $msg = $Btn2ID
				$Ok = 0
				Exitloop			
			Case $msg = $GUI_EVENT_CLOSE
				_Exit()
				Exitloop
		EndSelect
	Until $msg = $GUI_EVENT_CLOSE
	GUIDelete($dWnd)
	Return $Ok
EndFunc 